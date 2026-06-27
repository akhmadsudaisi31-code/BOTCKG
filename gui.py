import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import asyncio
import io
import os
import sys
from pathlib import Path
import pandas as pd
from PIL import Image, ImageTk

# Import the core bot script
import bot_kemkes

class BotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Arzachel Bot CKG (ASIK Kemenkes)")
        self.root.geometry("1000x700")
        self.root.minsize(950, 650)
        
        # Color Palette - Sleek Premium Dark Mode
        self.c_bg = "#121214"          # Deep dark background
        self.c_card = "#1a1a24"        # Card background
        self.c_card_elev = "#242432"   # Elevated card background
        self.c_text = "#e2e2e7"        # Bright text
        self.c_text_muted = "#8e8e9f"  # Muted grey text
        self.c_accent = "#00adb5"      # Cool teal accent
        self.c_accent_hover = "#008a90"
        
        self.c_green = "#2ecc71"       # Success
        self.c_green_hover = "#27ae60"
        self.c_yellow = "#f1c40f"      # Warning/Skip
        self.c_yellow_hover = "#d4af37"
        self.c_red = "#e74c3c"         # Danger/Fail
        self.c_red_hover = "#c0392b"
        self.c_blue = "#3498db"        # Info
        self.c_purple = "#9b59b6"      # Bot / Special
        
        self.root.configure(bg=self.c_bg)
        
        # Thread and queue management
        self.bot_thread = None
        self.bridge = None
        self.current_request_type = None
        self.current_request_data = None
        
        # App State
        self.is_running = False
        
        self.config_file = Path("gui_config.json")
        self.setup_styles()
        
        # Check activation status
        self.device_id = self.get_device_id()
        self.is_activated = self.check_activation()
        
        if not self.is_activated:
            self.build_activation_ui()
        else:
            self.build_ui()
            # Initial check for Excel
            self.auto_detect_excel()
            # Start periodic poller (50ms interval)
            self.root.after(50, self.poll_queues)
            # Check for updates in background (1 second delay)
            self.root.after(1000, self.check_for_updates)
        
    def load_config(self):
        import json
        defaults = {
            "email": bot_kemkes.USERNAME,
            "password": bot_kemkes.PASSWORD,
            "default_phone": "",
            "excel_path": str(bot_kemkes.FILE_EXCEL),
            "auto_advance": True,
            "headless": bot_kemkes.HEADLESS
        }
        if self.config_file.exists():
            try:
                with open(self.config_file, "r") as f:
                    data = json.load(f)
                    for k, v in data.items():
                        defaults[k] = v
            except Exception:
                pass
        return defaults

    def save_config(self):
        import json
        existing_data = {}
        if self.config_file.exists():
            try:
                with open(self.config_file, "r") as f:
                    existing_data = json.load(f)
            except Exception:
                pass
        data = {
            "email": self.get_entry_val(self.entry_email),
            "password": self.get_entry_val(self.entry_password),
            "default_phone": self.get_entry_val(self.entry_phone),
            "excel_path": self.get_entry_val(self.entry_excel),
            "auto_advance": self.var_auto_advance.get(),
            "headless": self.var_headless.get(),
            "activation_key": existing_data.get("activation_key", "")
        }
        try:
            with open(self.config_file, "w") as f:
                json.dump(data, f, indent=4)
        except Exception:
            pass

    def manual_save_config(self):
        self.save_config()
        messagebox.showinfo("Sukses", "Konfigurasi berhasil disimpan ke gui_config.json!")
        self.append_log("✓ Konfigurasi berhasil disimpan ke gui_config.json", "OK")

    def setup_styles(self):
        # Configure scrollbar style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background=self.c_bg, foreground=self.c_text)
        style.configure("TProgressbar", thickness=15, troughcolor=self.c_card, background=self.c_accent, bordercolor=self.c_bg)
        
        # Style Combobox for dark theme (ensuring text is visible)
        style.configure("TCombobox", 
                        fieldbackground=self.c_bg, 
                        background=self.c_card_elev, 
                        foreground=self.c_text, 
                        arrowcolor=self.c_text,
                        bordercolor="#323246",
                        lightcolor="#323246",
                        darkcolor="#323246")
        style.map("TCombobox", 
                  fieldbackground=[("readonly", self.c_bg), ("disabled", self.c_bg)], 
                  foreground=[("readonly", self.c_text), ("disabled", self.c_text_muted)],
                  selectbackground=[("readonly", self.c_bg)],
                  selectforeground=[("readonly", self.c_text)])
                  
        # Style the combobox dropdown popdown listbox to be dark themed
        self.root.option_add("*TCombobox*Listbox.background", self.c_card)
        self.root.option_add("*TCombobox*Listbox.foreground", self.c_text)
        self.root.option_add("*TCombobox*Listbox.selectBackground", self.c_accent)
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        self.root.option_add("*TCombobox*Listbox.font", ("Helvetica", 9))
        
    def build_ui(self):
        config_data = self.load_config()
        
        # Top container for banners/notices (e.g. Update Banner)
        self.banner_container = tk.Frame(self.root, bg=self.c_bg)
        self.banner_container.pack(side="top", fill="x")
        
        # Main layout: Two main columns
        self.main_pane = tk.Frame(self.root, bg=self.c_bg)
        self.main_pane.pack(fill="both", expand=True, padx=15, pady=15)
        
        self.left_col = tk.Frame(self.main_pane, bg=self.c_bg, width=420)
        self.left_col.pack(side="left", fill="both", expand=False, padx=(0, 10))
        self.left_col.pack_propagate(False)
        
        self.right_col = tk.Frame(self.main_pane, bg=self.c_bg)
        self.right_col.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # ------------------------------------------------------------
        #  LEFT COLUMN: CONFIGURATION & INTERACTION
        # ------------------------------------------------------------
        
        # Card 1: Configuration
        self.config_card = tk.Frame(self.left_col, bg=self.c_card, bd=0, highlightthickness=0)
        self.config_card.pack(fill="both", expand=False, pady=(0, 15))
        
        self.create_card_header(self.config_card, "⚙️ KONFIGURASI BOT")
        
        # Inner padding frame for config
        config_inner = tk.Frame(self.config_card, bg=self.c_card, padx=15, pady=10)
        config_inner.pack(fill="both")
        
        # Email & Password
        tk.Label(config_inner, text="Email / Username", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).grid(row=0, column=0, sticky="w", pady=(5, 2))
        self.entry_email = self.create_flat_entry(config_inner, config_data["email"])
        self.entry_email.container.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        tk.Label(config_inner, text="Kata Sandi (Password)", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(5, 2))
        self.entry_password = self.create_flat_entry(config_inner, config_data["password"], show="*")
        self.entry_password.container.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        # Nomor WhatsApp/HP Default
        tk.Label(config_inner, text="Nomor WhatsApp/HP Default", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).grid(row=4, column=0, sticky="w", pady=(5, 2))
        self.entry_phone = self.create_flat_entry(config_inner, config_data.get("default_phone", ""))
        self.entry_phone.container.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        # Excel File selection
        tk.Label(config_inner, text="File Excel Data Pasien (.xlsx)", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).grid(row=6, column=0, sticky="w", pady=(5, 2))
        self.entry_excel = self.create_flat_entry(config_inner, config_data["excel_path"])
        self.entry_excel.container.grid(row=7, column=0, sticky="ew", pady=(0, 10))
        
        self.btn_browse = self.create_flat_button(config_inner, "Cari...", self.browse_excel, bg=self.c_card_elev, hover_bg="#323246", width=8)
        self.btn_browse.grid(row=7, column=1, padx=(5, 0), pady=(0, 10), sticky="ns")
        
        # Starting row
        tk.Label(config_inner, text="Mulai dari No. Pasien di Excel (Kolom 'No')", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).grid(row=8, column=0, sticky="w", pady=(5, 2))
        self.entry_start_row = self.create_flat_entry(config_inner, "1")
        self.entry_start_row.container.grid(row=9, column=0, sticky="ew", pady=(0, 10))
        
        self.btn_detect = self.create_flat_button(config_inner, "Deteksi Lanjutan", self.auto_detect_excel, bg=self.c_card_elev, hover_bg="#323246", width=15)
        self.btn_detect.grid(row=9, column=1, padx=(5, 0), pady=(0, 10), sticky="ns")
        
        # Checkboxes for options
        self.var_auto_advance = tk.BooleanVar(value=config_data["auto_advance"])
        self.cb_auto_advance = tk.Checkbutton(
            config_inner, text="Maju Otomatis ke Pasien Berikutnya", variable=self.var_auto_advance,
            bg=self.c_card, fg=self.c_text, activebackground=self.c_card, activeforeground=self.c_text,
            selectcolor=self.c_bg, font=("Helvetica", 9)
        )
        self.cb_auto_advance.grid(row=10, column=0, columnspan=2, sticky="w", pady=(5, 2))
        
        self.var_headless = tk.BooleanVar(value=config_data["headless"])
        self.cb_headless = tk.Checkbutton(
            config_inner, text="Headless Mode (Sembunyikan Jendela Browser)", variable=self.var_headless,
            bg=self.c_card, fg=self.c_text, activebackground=self.c_card, activeforeground=self.c_text,
            selectcolor=self.c_bg, font=("Helvetica", 9)
        )
        self.cb_headless.grid(row=11, column=0, columnspan=2, sticky="w", pady=(2, 10))
        
        # Large Action Button
        self.btn_start = self.create_flat_button(
            config_inner, "🚀 JALANKAN BOT", self.toggle_bot, 
            bg=self.c_green, hover_bg=self.c_green_hover, height=2, font=("Helvetica", 11, "bold")
        )
        self.btn_start.grid(row=12, column=0, columnspan=2, sticky="ew", pady=(10, 5))
        
        # Simpan Konfigurasi Button
        self.btn_save = self.create_flat_button(
            config_inner, "💾 SIMPAN KONFIGURASI", self.manual_save_config, 
            bg=self.c_card_elev, hover_bg="#323246", height=1, font=("Helvetica", 9, "bold")
        )
        self.btn_save.grid(row=13, column=0, columnspan=2, sticky="ew", pady=(5, 5))
        
        # Card 2: Interactive Control Panel
        self.interact_card = tk.Frame(self.left_col, bg=self.c_card, bd=0, highlightthickness=0)
        self.interact_card.pack(fill="both", expand=True)
        
        # Copyright Label at the bottom of left column
        self.lbl_copyright = tk.Label(
            self.left_col, 
            text="© 2026 Arzachel Bot CKG. Hak Cipta Dilindungi Undang-Undang.", 
            bg=self.c_bg, fg=self.c_text_muted, font=("Helvetica", 8, "italic")
        )
        self.lbl_copyright.pack(side="bottom", fill="x", pady=(10, 0))
        
        self.create_card_header(self.interact_card, "🕹️ PANEL INTERAKSI / INPUT")
        
        self.interact_inner = tk.Frame(self.interact_card, bg=self.c_card, padx=15, pady=15)
        self.interact_inner.pack(fill="both", expand=True)
        
        self.show_default_interact_message()
        
        # ------------------------------------------------------------
        #  RIGHT COLUMN: PROGRESS, DATA PREVIEW & LOGS
        # ------------------------------------------------------------
        
        # Card 3: Progress & Patient Status
        self.progress_card = tk.Frame(self.right_col, bg=self.c_card, bd=0, highlightthickness=0)
        self.progress_card.pack(fill="x", expand=False, pady=(0, 15))
        
        self.create_card_header(self.progress_card, "📊 STATUS & PROGRES")
        
        prog_inner = tk.Frame(self.progress_card, bg=self.c_card, padx=15, pady=15)
        prog_inner.pack(fill="x")
        
        # Baris progres utama
        self.lbl_progress = tk.Label(prog_inner, text="Belum berjalan. Silakan klik tombol Jalankan Bot.", bg=self.c_card, fg=self.c_text, font=("Helvetica", 10, "bold"))
        self.lbl_progress.pack(anchor="w", pady=(0, 8))
        
        # Progress Bar
        self.progress_bar = ttk.Progressbar(prog_inner, orient="horizontal", mode="determinate", style="TProgressbar")
        self.progress_bar.pack(fill="x", pady=(0, 12))
        
        # Separator tipis
        tk.Frame(prog_inner, bg="#2a2a3a", height=1).pack(fill="x", pady=(0, 10))
        
        # Kartu info pasien aktif
        self.details_frame = tk.Frame(prog_inner, bg=self.c_card)
        self.details_frame.pack(fill="x")
        self.details_frame.columnconfigure(1, weight=1)

        # Label judul
        tk.Label(self.details_frame, text="Pasien Aktif:", bg=self.c_card,
                 fg=self.c_text_muted, font=("Helvetica", 8)).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        # Nama — lebih besar dan berwarna
        tk.Label(self.details_frame, text="Nama", bg=self.c_card,
                 fg=self.c_text_muted, font=("Helvetica", 8)).grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.lbl_patient_name = tk.Label(self.details_frame, text="—", bg=self.c_card,
                                          fg=self.c_accent, font=("Helvetica", 11, "bold"), anchor="w")
        self.lbl_patient_name.grid(row=1, column=1, sticky="ew", pady=(0, 3))

        # NIK
        tk.Label(self.details_frame, text="NIK  ", bg=self.c_card,
                 fg=self.c_text_muted, font=("Helvetica", 8)).grid(row=2, column=0, sticky="w", padx=(0, 8))
        self.lbl_patient_nik = tk.Label(self.details_frame, text="—", bg=self.c_card,
                                         fg="#f1c40f", font=("Helvetica", 10, "bold"), anchor="w")
        self.lbl_patient_nik.grid(row=2, column=1, sticky="ew")
        
        # Card 4: Log Console
        self.log_card = tk.Frame(self.right_col, bg=self.c_card, bd=0, highlightthickness=0)
        self.log_card.pack(fill="both", expand=True)
        
        self.create_card_header(self.log_card, "💻 LOG KONSOL")
        
        log_inner = tk.Frame(self.log_card, bg=self.c_card, padx=10, pady=10)
        log_inner.pack(fill="both", expand=True)
        
        # ScrolledText for logs
        self.log_text = scrolledtext.ScrolledText(
            log_inner, wrap="word", bg="#0a0a0f", fg="#cccccc",
            insertbackground="#ffffff", font=("Courier", 10), bd=0, highlightthickness=0
        )
        self.log_text.pack(fill="both", expand=True)
        
        # Configure tags for rich colors
        self.log_text.tag_config("INFO", foreground=self.c_blue)
        self.log_text.tag_config("OK", foreground=self.c_green)
        self.log_text.tag_config("WARN", foreground=self.c_yellow)
        self.log_text.tag_config("ERR", foreground=self.c_red)
        self.log_text.tag_config("BOT", foreground=self.c_accent)
        self.log_text.tag_config("WAIT", foreground=self.c_purple)
        self.log_text.tag_config("NORMAL", foreground="#cccccc")
        
        self.append_log("Sistem GUI siap. Silakan klik 'Jalankan Bot' untuk memulai.", "INFO")

    # ------------------------------------------------------------
    #  UI HELPER METHODS
    # ------------------------------------------------------------
    
    def create_card_header(self, parent, title):
        header = tk.Frame(parent, bg=self.c_card_elev, height=35)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        lbl = tk.Label(header, text=title, bg=self.c_card_elev, fg=self.c_text, font=("Helvetica", 9, "bold"))
        lbl.pack(side="left", padx=12, pady=8)
        
        # A tiny accent line at the bottom of the header
        accent_line = tk.Frame(parent, bg=self.c_accent, height=2)
        accent_line.pack(fill="x")

    def create_flat_entry(self, parent, default_text="", show=None):
        frame = tk.Frame(parent, bg=self.c_bg, bd=0, highlightthickness=1, highlightbackground="#323246", highlightcolor=self.c_accent)
        entry = tk.Entry(frame, bg=self.c_bg, fg=self.c_text, insertbackground=self.c_text, bd=0, font=("Helvetica", 10), show=show)
        entry.insert(0, default_text)
        entry.pack(fill="both", expand=True, padx=8, pady=5)
        entry.container = frame
        return entry

    def create_flat_button(self, parent, text, command, bg, fg="#ffffff", hover_bg=None, active_bg=None, width=15, height=1, font=("Helvetica", 9, "bold")):
        btn = tk.Button(
            parent, text=text, command=command, bg=bg, fg=fg, 
            activebackground=active_bg or hover_bg or bg, activeforeground=fg,
            bd=0, highlightthickness=0, relief="flat", font=font, 
            width=width, height=height, cursor="hand2"
        )
        if hover_bg:
            btn.bind("<Enter>", lambda e: btn.config(bg=hover_bg))
            btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    def get_entry_val(self, entry_widget):
        # Helper to get value of an entry inside our custom frame
        # If the entry_widget itself is passed, it is pack'ed inside a Frame
        # Let's see: self.entry_email is an Entry inside a Frame
        return entry_widget.get().strip()

    # ------------------------------------------------------------
    #  INTERACTION MGR - PANEL STATES
    # ------------------------------------------------------------
    
    def clear_interact_panel(self):
        for widget in self.interact_inner.winfo_children():
            widget.destroy()

    def show_default_interact_message(self):
        self.clear_interact_panel()
        
        lbl = tk.Label(
            self.interact_inner, 
            text="☕ Menunggu Aktivitas Bot...\n\nBot berjalan secara otomatis. Jika diperlukan tindakan manual Anda (seperti memasukkan CAPTCHA atau memilih status pasien), panel ini akan memandu Anda.",
            bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 10), justify="center", wraplength=350
        )
        lbl.pack(fill="both", expand=True, pady=40)

    def handle_request(self, req_type, data):
        self.current_request_type = req_type
        self.current_request_data = data
        self.clear_interact_panel()
        
        # Focus window to alert user
        self.root.attributes('-topmost', True)
        self.root.attributes('-topmost', False)
        
        if req_type == "captcha":
            self.build_captcha_ui(data)
        elif req_type == "decision":
            self.build_decision_ui(data)
        elif req_type == "wali":
            self.build_wali_ui(data)
        elif req_type == "manual_address":
            self.build_manual_address_ui(data)
        elif req_type == "kategori_pasien":
            self.build_kategori_pasien_ui(data)
        elif req_type == "login_failure":
            self.build_login_failure_ui(data)
        elif req_type == "next_patient_pause":
            self.build_next_patient_pause_ui(data)
        elif req_type == "take_control":
            self.build_take_control_ui(data)
        else:
            # Fallback
            self.show_default_interact_message()
            self.bridge.response_queue.put("continue")

    def send_response(self, response_val):
        self.current_request_type = None
        self.current_request_data = None
        self.show_default_interact_message()
        self.bridge.response_queue.put(response_val)

    # 1. CAPTCHA Panel
    def build_captcha_ui(self, img_bytes):
        tk.Label(self.interact_inner, text="🔒 VERIFIKASI CAPTCHA", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        tk.Label(self.interact_inner, text="Masukkan kode CAPTCHA yang muncul di bawah ini untuk melanjutkan login.", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 15))
        
        # Display Image
        img_label = tk.Label(self.interact_inner, bg=self.c_bg, bd=1, relief="solid")
        img_label.pack(pady=(0, 15), ipady=5, ipadx=5)
        
        if img_bytes:
            try:
                img = Image.open(io.BytesIO(img_bytes))
                # Scale up to make it highly readable
                scaled_w = img.width * 2.2
                scaled_h = img.height * 2.2
                img = img.resize((int(scaled_w), int(scaled_h)), Image.Resampling.LANCZOS)
                img_tk = ImageTk.PhotoImage(img)
                img_label.image = img_tk
                img_label.config(image=img_tk)
            except Exception as e:
                img_label.config(text=f"Gagal memuat Gambar: {e}", fg=self.c_red)
        else:
            img_label.config(text="Gambar CAPTCHA tidak ditemukan.\nPeriksa layar browser!", fg=self.c_yellow, padx=20, pady=10)
            
        # Text input
        tk.Label(self.interact_inner, text="Kode CAPTCHA", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(0, 2))
        
        entry_frame = tk.Frame(self.interact_inner, bg=self.c_bg, bd=0, highlightthickness=1, highlightbackground="#323246", highlightcolor=self.c_accent)
        entry_frame.pack(fill="x", pady=(0, 15))
        
        captcha_entry = tk.Entry(entry_frame, bg=self.c_bg, fg=self.c_text, insertbackground=self.c_text, bd=0, font=("Helvetica", 12, "bold"), justify="center")
        captcha_entry.pack(fill="both", expand=True, padx=8, pady=6)
        captcha_entry.focus_set()
        
        # Button submit
        btn_submit = self.create_flat_button(
            self.interact_inner, "SUBMIT CAPTCHA", 
            lambda: self.send_response(captcha_entry.get().strip()), 
            bg=self.c_accent, hover_bg=self.c_accent_hover, height=2, font=("Helvetica", 10, "bold")
        )
        btn_submit.pack(fill="x")
        
        # Bind Enter key to submit
        captcha_entry.bind("<Return>", lambda e: self.send_response(captcha_entry.get().strip()))

    # 2. Excel Row Decision Panel
    def build_decision_ui(self, data):
        nomor = data["nomor"]
        nama = data["nama"]
        data_tidak_valid = data["data_tidak_valid"]
        default_pilihan = data["default_pilihan"]
        
        tk.Label(self.interact_inner, text="📋 TENTUKAN STATUS EXCEL", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        
        txt_desc = f"Form untuk Pasien No. {nomor} ({nama}) telah selesai diisi.\nTentukan status yang akan ditulis di file Excel:"
        tk.Label(self.interact_inner, text=txt_desc, bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 10))
        
        if data_tidak_valid:
            warn_frame = tk.Frame(self.interact_inner, bg="#2c1a1a", bd=1, relief="solid", highlightthickness=0)
            warn_frame.pack(fill="x", pady=(0, 15), ipady=8, ipadx=8)
            tk.Label(
                warn_frame, 
                text="⚠️ PERINGATAN: Portal Kemkes melaporkan data pasien ini TIDAK VALID (misal: NIK salah / tidak terdaftar). Direkomendasikan memilih SKIP.", 
                bg="#2c1a1a", fg=self.c_red, font=("Helvetica", 8, "bold"), wraplength=330, justify="left"
            ).pack(fill="x")
            
        # Buttons grid
        btn_frame = tk.Frame(self.interact_inner, bg=self.c_card)
        btn_frame.pack(fill="x", pady=5)
        
        # 1. Lolos (Success)
        btn_lolos = self.create_flat_button(
            btn_frame, "🟢 1. LOLOS (Hijau)", 
            lambda: self.send_response("1"), 
            bg=self.c_green, hover_bg=self.c_green_hover, height=2, font=("Helvetica", 9, "bold")
        )
        btn_lolos.pack(fill="x", pady=(0, 8))
        
        # 2. Skip (Warning)
        btn_skip = self.create_flat_button(
            btn_frame, "🟡 2. SKIP (Kuning)", 
            lambda: self.send_response("2"), 
            bg=self.c_yellow, fg="#000000", hover_bg=self.c_yellow_hover, height=2, font=("Helvetica", 9, "bold")
        )
        btn_skip.pack(fill="x", pady=(0, 8))
        
        # 3. Gagal (Error)
        btn_gagal = self.create_flat_button(
            btn_frame, "🔴 3. GAGAL (Merah)", 
            lambda: self.send_response("3"), 
            bg=self.c_red, hover_bg=self.c_red_hover, height=2, font=("Helvetica", 9, "bold")
        )
        btn_gagal.pack(fill="x", pady=(0, 15))
        
        # Options row: Repeat / Quit
        opt_frame = tk.Frame(self.interact_inner, bg=self.c_card)
        opt_frame.pack(fill="x")
        
        btn_ulang = self.create_flat_button(
            opt_frame, "🔄 Ulangi Baris Ini", 
            lambda: self.send_response("r"), 
            bg=self.c_card_elev, hover_bg="#323246", width=18, height=1
        )
        btn_ulang.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        btn_berhenti = self.create_flat_button(
            opt_frame, "🛑 Berhenti (Quit)", 
            lambda: self.send_response("q"), 
            bg="#2c2c35", hover_bg="#1c1c22", width=18, height=1
        )
        btn_berhenti.pack(side="right", expand=True, fill="x", padx=(5, 0))
        
        # Focus on the recommended button
        if default_pilihan == "1":
            btn_lolos.focus_set()
        else:
            btn_skip.focus_set()

    # 3. Guardian (Wali) Selection Panel
    def build_wali_ui(self, data):
        no_excel = data["no_excel"]
        nama = data["nama"]
        above_rows = data["above_rows"]
        
        tk.Label(self.interact_inner, text="👨‍👩‍👦 PILIH DATA WALI", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        
        txt_desc = f"Pasien No. {no_excel} ({nama}) terdeteksi sebagai anak-anak/memerlukan Wali.\nPilih baris di Excel untuk dijadikan Wali:"
        tk.Label(self.interact_inner, text=txt_desc, bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 10))
        
        # Dropdown listbox / combobox of above rows
        tk.Label(self.interact_inner, text="Baris Pasien di Atas (Rekomendasi)", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 8, "bold")).pack(anchor="w", pady=(0, 2))
        
        combo_frame = tk.Frame(self.interact_inner, bg=self.c_bg, bd=0, highlightthickness=1, highlightbackground="#323246")
        combo_frame.pack(fill="x", pady=(0, 10))
        
        # Prepare choices: key is string display, value is "No"
        choices = ["(Gunakan Data Pasien Sendiri sebagai Wali)"]
        no_map = {"(Gunakan Data Pasien Sendiri sebagai Wali)": ""}
        
        for r in above_rows:
            display = f"No. {r['no']} | {r['nama']} (Baris: {r['excel_row']})"
            choices.append(display)
            no_map[display] = str(r["no"])
            
        combo_var = tk.StringVar(value=choices[0])
        combo = ttk.Combobox(combo_frame, textvariable=combo_var, values=choices, state="readonly", font=("Helvetica", 9))
        combo.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Manual Input
        tk.Label(self.interact_inner, text="Atau Masukkan Nomor Pasien Lain (Kolom 'No')", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 8, "bold")).pack(anchor="w", pady=(5, 2))
        entry_frame = tk.Frame(self.interact_inner, bg=self.c_bg, bd=0, highlightthickness=1, highlightbackground="#323246", highlightcolor=self.c_accent)
        entry_frame.pack(fill="x", pady=(0, 15))
        
        custom_entry = tk.Entry(entry_frame, bg=self.c_bg, fg=self.c_text, insertbackground=self.c_text, bd=0, font=("Helvetica", 9))
        custom_entry.pack(fill="both", expand=True, padx=8, pady=5)
        
        # Submit
        def on_submit():
            val = custom_entry.get().strip()
            if not val:
                selected = combo_var.get()
                val = no_map.get(selected, "")
            self.send_response(val)
            
        btn_submit = self.create_flat_button(
            self.interact_inner, "KIRIM DATA WALI", 
            on_submit, 
            bg=self.c_accent, hover_bg=self.c_accent_hover, height=2, font=("Helvetica", 10, "bold")
        )
        btn_submit.pack(fill="x")

    # 4. Manual Address Input Panel
    def build_manual_address_ui(self, data):
        prov = data.get("provinsi", "-")
        kab = data.get("kabupaten", "-")
        kec = data.get("kecamatan", "-")
        kel = data.get("kelurahan", "-")
        
        tk.Label(self.interact_inner, text="🏠 ALAMAT DOMISILI MANUAL", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        tk.Label(self.interact_inner, text="Bot memerlukan bantuan Anda untuk memilih alamat domisili di browser secara manual.", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 15))
        
        # Details box
        details_box = tk.Frame(self.interact_inner, bg=self.c_card_elev, padx=12, pady=10)
        details_box.pack(fill="x", pady=(0, 20))
        
        tk.Label(details_box, text=f"Provinsi       : {prov}", bg=self.c_card_elev, fg=self.c_text, font=("Courier", 9), anchor="w").pack(fill="x", pady=2)
        tk.Label(details_box, text=f"Kabupaten/Kota : {kab}", bg=self.c_card_elev, fg=self.c_text, font=("Courier", 9), anchor="w").pack(fill="x", pady=2)
        tk.Label(details_box, text=f"Kecamatan      : {kec}", bg=self.c_card_elev, fg=self.c_text, font=("Courier", 9), anchor="w").pack(fill="x", pady=2)
        tk.Label(details_box, text=f"Desa/Kelurahan : {kel}", bg=self.c_card_elev, fg=self.c_text, font=("Courier", 9), anchor="w").pack(fill="x", pady=2)
        
        instruction_frame = tk.Frame(self.interact_inner, bg="#2a2010", bd=1, relief="solid", highlightthickness=0)
        instruction_frame.pack(fill="x", pady=(0, 15), ipady=8, ipadx=8)
        tk.Label(
            instruction_frame, 
            text="⚠️ PETUNJUK: Pergi ke jendela browser Chromium, ketik dan cari alamat di atas pada form drop-down, lalu setelah selesai, klik tombol hijau di bawah ini.", 
            bg="#2a2010", fg=self.c_yellow, font=("Helvetica", 8, "bold"), wraplength=330, justify="left"
        ).pack(fill="x")
        
        btn_continue = self.create_flat_button(
            self.interact_inner, "✓ SAYA SUDAH ISI ALAMAT", 
            lambda: self.send_response("continue"), 
            bg=self.c_green, hover_bg=self.c_green_hover, height=2, font=("Helvetica", 10, "bold")
        )
        btn_continue.pack(fill="x")

    # 5. Kategori Pasien Warning Panel
    def build_kategori_pasien_ui(self, msg_text):
        tk.Label(self.interact_inner, text="⚠️ PERINGATAN KATEGORI SASARAN", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        
        tk.Label(self.interact_inner, text="Deteksi peringatan dari portal ASIK:", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9)).pack(anchor="w", pady=(0, 5))
        
        # Warning msg
        msg_frame = tk.Frame(self.interact_inner, bg="#2c1a1a", bd=1, relief="solid", highlightthickness=0)
        msg_frame.pack(fill="x", pady=(0, 15), ipady=10, ipadx=10)
        tk.Label(
            msg_frame, text=f"\"{msg_text}\"", 
            bg="#2c1a1a", fg=self.c_red, font=("Helvetica", 9, "italic"), wraplength=330, justify="left"
        ).pack(fill="x")
        
        tk.Label(self.interact_inner, text="Apakah Anda ingin tetap meloloskan pasien ini untuk didaftarkan, atau melewatinya (skip)?", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 20))
        
        btn_frame = tk.Frame(self.interact_inner, bg=self.c_card)
        btn_frame.pack(fill="x")
        
        btn_lolos = self.create_flat_button(
            btn_frame, "✓ LOLOSKAN PASIEN", 
            lambda: self.send_response("l"), 
            bg=self.c_green, hover_bg=self.c_green_hover, width=18, height=2
        )
        btn_lolos.pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        btn_skip = self.create_flat_button(
            btn_frame, "⏭ SKIP PASIEN", 
            lambda: self.send_response("s"), 
            bg=self.c_yellow, fg="#000000", hover_bg=self.c_yellow_hover, width=18, height=2
        )
        btn_skip.pack(side="right", expand=True, fill="x", padx=(5, 0))

    # 6. Login Failure Panel
    def build_login_failure_ui(self, current_url):
        tk.Label(self.interact_inner, text="⚠️ LOGIN DITOLAK / GAGAL", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        tk.Label(self.interact_inner, text=f"Browser gagal menavigasi ke Dashboard (URL saat ini: {current_url}).\nPilih tindakan yang akan diambil:", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 15))
        
        # Action Options
        # 1. Retry
        btn_retry = self.create_flat_button(
            self.interact_inner, "🔄 1. Coba Ulangi Login", 
            lambda: self.send_response("1"), 
            bg=self.c_accent, hover_bg=self.c_accent_hover, height=2, font=("Helvetica", 9, "bold")
        )
        btn_retry.pack(fill="x", pady=(0, 8))
        
        # 3. Force Continue
        btn_force = self.create_flat_button(
            self.interact_inner, "➡️ 3. Paksa Lanjut (Abaikan Dashboard)", 
            lambda: self.send_response("3"), 
            bg=self.c_card_elev, hover_bg="#323246", height=2, font=("Helvetica", 9, "bold")
        )
        btn_force.pack(fill="x", pady=(0, 15))
        
        # 2. Enter new credentials inline
        tk.Label(self.interact_inner, text="Atau Masukkan Akun Baru (Email & Sandi Baru):", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(5, 2))
        
        entry_frame_e = tk.Frame(self.interact_inner, bg=self.c_bg, bd=0, highlightthickness=1, highlightbackground="#323246")
        entry_frame_e.pack(fill="x", pady=(0, 8))
        new_email = tk.Entry(entry_frame_e, bg=self.c_bg, fg=self.c_text, insertbackground=self.c_text, bd=0, font=("Helvetica", 9))
        new_email.insert(0, "Email baru...")
        new_email.pack(fill="both", expand=True, padx=8, pady=5)
        
        entry_frame_p = tk.Frame(self.interact_inner, bg=self.c_bg, bd=0, highlightthickness=1, highlightbackground="#323246")
        entry_frame_p.pack(fill="x", pady=(0, 12))
        new_pass = tk.Entry(entry_frame_p, bg=self.c_bg, fg=self.c_text, insertbackground=self.c_text, bd=0, font=("Helvetica", 9), show="*")
        new_pass.insert(0, "Sandi baru...")
        new_pass.pack(fill="both", expand=True, padx=8, pady=5)
        
        btn_new_acc = self.create_flat_button(
            self.interact_inner, "🔑 2. Masuk dengan Akun Baru", 
            lambda: self.send_response(("2", new_email.get().strip(), new_pass.get().strip())), 
            bg=self.c_green, hover_bg=self.c_green_hover, height=2, font=("Helvetica", 9, "bold")
        )
        btn_new_acc.pack(fill="x")

    # 7. Next Patient Pause Panel (when auto_advance is False)
    def build_next_patient_pause_ui(self, nomor):
        tk.Label(self.interact_inner, text="⏸️ JEDA ANTAR PASIEN", bg=self.c_card, fg=self.c_text, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        
        txt_desc = f"Baris {nomor} selesai diproses.\nBot sedang dijeda sesuai opsi 'Jeda Manual'. Klik tombol di bawah untuk melanjutkan ke pasien berikutnya."
        tk.Label(self.interact_inner, text=txt_desc, bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 20))
        
        btn_continue = self.create_flat_button(
            self.interact_inner, "➡️ LANJUTKAN BOT", 
            lambda: self.send_response("continue"), 
            bg=self.c_green, hover_bg=self.c_green_hover, height=2, font=("Helvetica", 11, "bold")
        )
        btn_continue.pack(fill="x", pady=(0, 10))
        
        btn_quit = self.create_flat_button(
            self.interact_inner, "🛑 Berhenti & Keluar", 
            lambda: self.send_response("q"), 
            bg="#2c2c35", hover_bg="#1c1c22", height=1.5
        )
        btn_quit.pack(fill="x")
        
        btn_continue.focus_set()

    # 8. Smart Assistant (Take Control) Panel
    def build_take_control_ui(self, data):
        nomor = data["nomor"]
        nama = data["nama"]
        error_msg = data["error_msg"]
        
        tk.Label(self.interact_inner, text="⚠️ ASISTEN PINTAR: AMBIL KENDALI", bg=self.c_card, fg=self.c_yellow, font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 5))
        
        txt_desc = f"Bot mendeteksi kondisi di luar kendali saat memproses Pasien No. {nomor} ({nama})."
        tk.Label(self.interact_inner, text=txt_desc, bg=self.c_card, fg=self.c_text, font=("Helvetica", 9), wraplength=360, justify="left").pack(anchor="w", pady=(0, 10))
        
        # Display the exact error message
        err_frame = tk.Frame(self.interact_inner, bg="#2a2010", bd=1, relief="solid")
        err_frame.pack(fill="x", pady=(0, 10), ipady=6, ipadx=8)
        tk.Label(
            err_frame, text=f"Detail Kendala:\n\"{error_msg[:120]}...\"", 
            bg="#2a2010", fg=self.c_yellow, font=("Helvetica", 8, "italic"), wraplength=330, justify="left"
        ).pack(fill="x")
        
        instruction_frame = tk.Frame(self.interact_inner, bg=self.c_card_elev, bd=0)
        instruction_frame.pack(fill="x", pady=(0, 10), ipady=6, ipadx=8)
        tk.Label(
            instruction_frame, 
            text="👉 PETUNJUK: Silakan lihat browser Chromium Anda. Perbaiki masalah secara manual (misal: tutup popup, lengkapi kolom yang macet, atau submit sendiri).\n\nSetelah selesai, silakan pilih tindakan di bawah ini:", 
            bg=self.c_card_elev, fg=self.c_text_muted, font=("Helvetica", 8), wraplength=330, justify="left"
        ).pack(fill="x")
        
        # Grid frame for action buttons (2 columns)
        grid_frame = tk.Frame(self.interact_inner, bg=self.c_card)
        grid_frame.pack(fill="x", pady=5)
        
        # Row 0
        btn_resume = self.create_flat_button(
            grid_frame, "▶️ RESUME", 
            lambda: self.send_response("resume"), 
            bg=self.c_accent, hover_bg=self.c_accent_hover, width=17, height=2
        )
        btn_resume.grid(row=0, column=0, padx=(0, 4), pady=4, sticky="ew")
        
        btn_berhasil = self.create_flat_button(
            grid_frame, "🟢 TANDAI LOLOS", 
            lambda: self.send_response("ok"), 
            bg=self.c_green, hover_bg=self.c_green_hover, width=17, height=2
        )
        btn_berhasil.grid(row=0, column=1, padx=(4, 0), pady=4, sticky="ew")
        
        # Row 1
        btn_ulang = self.create_flat_button(
            grid_frame, "🔄 ULANGI BARIS", 
            lambda: self.send_response("ulang"), 
            bg="#d35400", hover_bg="#e67e22", width=17, height=2
        )
        btn_ulang.grid(row=1, column=0, padx=(0, 4), pady=4, sticky="ew")
        
        btn_skip = self.create_flat_button(
            grid_frame, "🟡 LEWATI (SKIP)", 
            lambda: self.send_response("skip"), 
            bg=self.c_yellow, fg="#000000", hover_bg=self.c_yellow_hover, width=17, height=2
        )
        btn_skip.grid(row=1, column=1, padx=(4, 0), pady=4, sticky="ew")
        
        # Row 2
        btn_gagal = self.create_flat_button(
            grid_frame, "🔴 TANDAI GAGAL", 
            lambda: self.send_response("gagal"), 
            bg=self.c_red, hover_bg=self.c_red_hover, width=17, height=2
        )
        btn_gagal.grid(row=2, column=0, padx=(0, 4), pady=4, sticky="ew")
        
        btn_quit = self.create_flat_button(
            grid_frame, "🛑 HENTIKAN BOT", 
            lambda: self.send_response("quit"), 
            bg="#2c2c35", hover_bg="#1c1c22", width=17, height=2
        )
        btn_quit.grid(row=2, column=1, padx=(4, 0), pady=4, sticky="ew")
        
        # Configure grid columns to have equal width
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)
        
        btn_resume.focus_set()
 
    # ------------------------------------------------------------
    #  LICENSING & ACTIVATION SYSTEM
    # ------------------------------------------------------------
    
    def get_device_id(self):
        import uuid
        import hashlib
        mac = str(uuid.getnode())
        h = hashlib.sha256(f"botckg-{mac}".encode()).hexdigest()
        return f"{h[0:4]}-{h[4:8]}-{h[8:12]}".upper()

    def generate_activation_key(self, device_id):
        import hashlib
        secret_salt = "arzachel_bot_ckg_salt_2026_super_secure"
        h = hashlib.sha256(f"{device_id}-{secret_salt}".encode()).hexdigest()
        return f"{h[12:16]}-{h[16:20]}-{h[20:24]}-{h[24:28]}".upper()

    def check_activation(self):
        config = self.load_config()
        saved_key = config.get("activation_key", "")
        expected_key = self.generate_activation_key(self.device_id)
        return saved_key == expected_key

    def build_activation_ui(self):
        for widget in self.root.winfo_children():
            widget.destroy()
            
        self.root.title("Aktivasi Arzachel Bot CKG")
        
        # Center container
        container = tk.Frame(self.root, bg=self.c_bg)
        container.place(relx=0.5, rely=0.5, anchor="center")
        
        # Premium Card
        card = tk.Frame(container, bg=self.c_card, highlightthickness=1, highlightbackground=self.c_accent, padx=30, pady=30)
        card.pack()
        
        # Lock Icon & Title
        lbl_lock = tk.Label(card, text="🔒", bg=self.c_card, fg=self.c_accent, font=("Helvetica", 32))
        lbl_lock.pack(pady=(0, 10))
        
        lbl_title = tk.Label(card, text="AKTIVASI APLIKASI", bg=self.c_card, fg=self.c_text, font=("Helvetica", 14, "bold"))
        lbl_title.pack(pady=(0, 5))
        
        lbl_sub = tk.Label(card, text="Arzachel Bot CKG (ASIK Kemenkes) memerlukan lisensi untuk digunakan.", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 9))
        lbl_sub.pack(pady=(0, 20))
        
        # Payment Info Box
        pay_info_text = (
            "💳 METODE AKTIVASI (Rp 160.000):\n"
            "- Transfer Bank: Bank Jago\n"
            "  Nomor Rekening: 105295129701\n"
            "  Atas Nama: Akhmad Sudaisi\n\n"
            "Langkah Aktivasi:\n"
            "1. Lakukan transfer sebesar Rp 160.000\n"
            "2. Salin DEVICE ID di bawah ini\n"
            "3. Kirimkan bukti transfer + DEVICE ID ke WhatsApp Admin\n"
            "4. Masukkan KUNCI AKTIVASI yang Anda terima di bawah"
        )
        pay_box = tk.Text(card, bg="#0a0a0f", fg="#cccccc", font=("Helvetica", 9), wrap="word", width=50, height=8, bd=0, highlightthickness=1, highlightbackground="#323246", padx=10, pady=10)
        pay_box.insert("1.0", pay_info_text)
        pay_box.config(state="disabled")
        pay_box.pack(pady=(0, 15))
        
        # QRIS / Barcode scan logic if local barcode image exists
        qris_path = None
        for name in ["qris", "barcode", "QRIS", "BARCODE"]:
            for ext in [".png", ".jpg", ".jpeg", ".PNG", ".JPG", ".JPEG"]:
                p = Path(f"{name}{ext}")
                if p.exists():
                    qris_path = p
                    break
            if qris_path:
                break
                
        if qris_path:
            try:
                img_qris = Image.open(qris_path)
                # Fit to 260x260 pixels to make it larger and easier to scan
                img_qris = img_qris.resize((260, 260), Image.Resampling.LANCZOS)
                img_tk = ImageTk.PhotoImage(img_qris)
                
                qris_frame = tk.Frame(card, bg=self.c_bg, bd=1, relief="solid", highlightthickness=0)
                qris_frame.pack(pady=(0, 10))
                
                lbl_qris_img = tk.Label(qris_frame, image=img_tk, bg=self.c_bg)
                lbl_qris_img.image = img_tk  # Keep reference
                lbl_qris_img.pack(padx=5, pady=5)
                
                lbl_qris_hint = tk.Label(card, text="📸 Scan Barcode untuk Membayar Rp 160.000", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 8, "italic"))
                lbl_qris_hint.pack(pady=(0, 15))
            except Exception as e:
                lbl_err = tk.Label(card, text=f"Gagal memuat barcode: {e}", bg=self.c_card, fg=self.c_red, font=("Helvetica", 8))
                lbl_err.pack(pady=(0, 10))
        
        # Account Number Box with Copy Button
        bank_frame = tk.Frame(card, bg=self.c_bg, highlightthickness=1, highlightbackground="#323246")
        bank_frame.pack(fill="x", pady=(0, 15))
        
        bank_details = tk.Frame(bank_frame, bg=self.c_bg)
        bank_details.pack(side="left", padx=10, pady=5)
        
        tk.Label(bank_details, text="BANK JAGO (Akhmad Sudaisi)", bg=self.c_bg, fg=self.c_text_muted, font=("Helvetica", 8, "bold"), anchor="w").pack(fill="x")
        lbl_rek = tk.Label(bank_details, text="105295129701", bg=self.c_bg, fg=self.c_green, font=("Courier", 12, "bold"), anchor="w")
        lbl_rek.pack(fill="x")
        
        def copy_rek():
            self.root.clipboard_clear()
            self.root.clipboard_append("105295129701")
            messagebox.showinfo("Sukses", "Nomor Rekening Bank Jago berhasil disalin ke clipboard!")
            
        btn_copy_rek = self.create_flat_button(bank_frame, "📋 SALIN REK", copy_rek, bg=self.c_card_elev, hover_bg="#323246", width=12)
        btn_copy_rek.pack(side="right", padx=10, pady=5)
        
        # Device ID
        device_frame = tk.Frame(card, bg=self.c_card)
        device_frame.pack(fill="x", pady=(0, 15))
        
        tk.Label(device_frame, text="DEVICE ID ANDA:", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 8, "bold")).pack(anchor="w")
        
        id_inner = tk.Frame(device_frame, bg=self.c_bg, highlightthickness=1, highlightbackground="#323246")
        id_inner.pack(fill="x", pady=(2, 0))
        
        lbl_device_id = tk.Label(id_inner, text=self.device_id, bg=self.c_bg, fg=self.c_yellow, font=("Courier", 12, "bold"))
        lbl_device_id.pack(side="left", padx=10, pady=5)
        
        # Copy Button
        def copy_id():
            self.root.clipboard_clear()
            self.root.clipboard_append(self.device_id)
            messagebox.showinfo("Sukses", "Device ID berhasil disalin ke clipboard!")
            
        btn_copy = self.create_flat_button(id_inner, "📋 SALIN", copy_id, bg=self.c_card_elev, hover_bg="#323246", width=8)
        btn_copy.pack(side="right", padx=5, pady=3)
        
        # Contact Admin Button
        def open_wa():
            import webbrowser
            url = f"https://wa.me/6282333017615?text=Halo%20Admin,%20saya%20ingin%20aktivasi%20Bot%20CKG.%20Ini%20Device%20ID%20saya:%20{self.device_id}"
            webbrowser.open(url)
            
        btn_wa = self.create_flat_button(card, "💬 HUBUNGI ADMIN VIA WHATSAPP", open_wa, bg="#25d366", hover_bg="#128c7e", height=1)
        btn_wa.pack(fill="x", pady=(0, 20))
        
        # Activation Key Input
        tk.Label(card, text="KUNCI AKTIVASI:", bg=self.c_card, fg=self.c_text_muted, font=("Helvetica", 8, "bold")).pack(anchor="w")
        
        key_inner = tk.Frame(card, bg=self.c_bg, highlightthickness=1, highlightbackground="#323246", highlightcolor=self.c_accent)
        key_inner.pack(fill="x", pady=(2, 15))
        
        self.entry_activation_key = tk.Entry(key_inner, bg=self.c_bg, fg=self.c_text, insertbackground=self.c_text, bd=0, font=("Helvetica", 11, "bold"), justify="center")
        self.entry_activation_key.pack(fill="both", expand=True, padx=8, pady=6)
        self.entry_activation_key.focus_set()
        
        # Activate Button
        def on_activate():
            key = self.entry_activation_key.get().strip()
            if not key:
                messagebox.showerror("Error", "Silakan masukkan Kunci Aktivasi!")
                return
                
            expected = self.generate_activation_key(self.device_id)
            if key.upper() == expected:
                # Save key to config
                import json
                config = {}
                if self.config_file.exists():
                    try:
                        with open(self.config_file, "r") as f:
                            config = json.load(f)
                    except Exception:
                        pass
                config["activation_key"] = key.upper()
                try:
                    with open(self.config_file, "w") as f:
                        json.dump(config, f, indent=4)
                except Exception:
                    pass
                    
                messagebox.showinfo("Sukses", "Aktivasi Berhasil! Terima kasih telah membeli lisensi.")
                
                # Rebuild main UI
                for widget in self.root.winfo_children():
                    widget.destroy()
                self.build_ui()
                self.auto_detect_excel()
                self.root.after(50, self.poll_queues)
                self.root.after(1000, self.check_for_updates)
            else:
                messagebox.showerror("Gagal Aktivasi", "Kunci Aktivasi tidak valid! Silakan periksa kembali atau hubungi admin.")
                
        btn_activate = self.create_flat_button(card, "✅ AKTIFKAN SEKARANG", on_activate, bg=self.c_green, hover_bg=self.c_green_hover, height=2, font=("Helvetica", 10, "bold"))
        btn_activate.pack(fill="x")
        
        # Bind Return key to activation
        self.entry_activation_key.bind("<Return>", lambda e: on_activate())

    # ------------------------------------------------------------
    #  AUTO-UPDATE SYSTEM
    # ------------------------------------------------------------
    
    def check_for_updates(self):
        self.append_log("Memeriksa pembaruan bot...", "INFO")
        threading.Thread(target=self._bg_check_updates, daemon=True).start()

    def _bg_check_updates(self):
        import subprocess
        
        # Configure Git to not prompt for credentials if not configured (to avoid hanging)
        env = os.environ.copy()
        env["GIT_TERMINAL_PROMPT"] = "0"
        env["GIT_SSH_COMMAND"] = "ssh -o BatchMode=yes"
        
        try:
            # 1. Run git fetch to update remote tracking branch info
            subprocess.run(["git", "fetch"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env, timeout=15)
            
            # 2. Get local commit hash
            res_local = subprocess.run(["git", "rev-parse", "HEAD"], capture_output=True, text=True, env=env, timeout=5)
            local_hash = res_local.stdout.strip()
            
            # 3. Get remote tracking branch commit hash
            res_remote = subprocess.run(["git", "rev-parse", "@{u}"], capture_output=True, text=True, env=env, timeout=5)
            remote_hash = res_remote.stdout.strip()
            
            if local_hash and remote_hash and local_hash != remote_hash:
                # There is an update! Show banner
                self.root.after(0, self.show_update_banner)
            else:
                self.root.after(0, lambda: self.append_log("Aplikasi sudah menggunakan versi terbaru.", "OK"))
        except Exception as e:
            # Silent fallback if offline or git repo hasn't been set up with remote yet
            self.root.after(0, lambda: self.append_log("Pemeriksaan update selesai (remote repository tidak aktif/belum dikonfigurasi).", "INFO"))

    def show_update_banner(self):
        # Clear existing banner widgets if any
        for widget in self.banner_container.winfo_children():
            widget.destroy()
            
        self.append_log("Pembaruan aplikasi terdeteksi di server!", "WARN")
        
        # Premium Slate Blue banner with Cyan/Teal border
        banner = tk.Frame(self.banner_container, bg="#1e293b", highlightthickness=1, highlightbackground=self.c_accent, height=45)
        banner.pack(fill="x", padx=15, pady=(15, 0))
        banner.pack_propagate(False)
        
        # Message Label
        lbl = tk.Label(
            banner, text="📢 Pembaruan Tersedia! Silakan unduh versi terbaru untuk performa dan fitur terbaru.", 
            bg="#1e293b", fg=self.c_text, font=("Helvetica", 9, "bold")
        )
        lbl.pack(side="left", padx=15, pady=10)
        
        # Update Button
        self.btn_run_update = self.create_flat_button(
            banner, "📥 PERBARUI SEKARANG", self.run_application_update,
            bg=self.c_accent, hover_bg=self.c_accent_hover, width=20, font=("Helvetica", 8, "bold")
        )
        self.btn_run_update.pack(side="right", padx=10, pady=8)
        
        # Close Button
        btn_close = tk.Button(
            banner, text="✕", command=lambda: banner.destroy(),
            bg="#1e293b", fg=self.c_text_muted, activebackground="#1e293b", activeforeground=self.c_text,
            bd=0, highlightthickness=0, relief="flat", font=("Helvetica", 10, "bold"), cursor="hand2"
        )
        btn_close.pack(side="right", padx=(5, 10), pady=10)
        btn_close.bind("<Enter>", lambda e: btn_close.config(fg=self.c_text))
        btn_close.bind("<Leave>", lambda e: btn_close.config(fg=self.c_text_muted))

    def run_application_update(self):
        self.btn_run_update.config(state="disabled", text="⏳ MEMPERBARUI...")
        self.append_log("Memulai proses pembaruan otomatis di latar belakang...", "WAIT")
        
        def do_update():
            import subprocess
            import sys
            
            env = os.environ.copy()
            env["GIT_TERMINAL_PROMPT"] = "0"
            env["GIT_SSH_COMMAND"] = "ssh -o BatchMode=yes"
            
            try:
                # 1. git pull
                self.root.after(0, lambda: self.append_log("Menjalankan git pull...", "WAIT"))
                res_pull = subprocess.run(["git", "pull"], capture_output=True, text=True, env=env, timeout=30)
                
                if res_pull.returncode != 0:
                    err_msg = res_pull.stderr.strip() or res_pull.stdout.strip()
                    raise Exception(f"Gagal melakukan git pull: {err_msg}")
                    
                self.root.after(0, lambda: self.append_log("Git pull berhasil dijalankan.", "OK"))
                if res_pull.stdout:
                    self.root.after(0, lambda: self.append_log(res_pull.stdout.strip(), "INFO"))
                
                # Get the correct paths of pip and playwright relative to sys.executable (inside venv)
                python_dir = Path(sys.executable).parent
                if sys.platform == "win32":
                    pip_path = str(python_dir / "pip.exe")
                    playwright_path = str(python_dir / "playwright.exe")
                else:
                    pip_path = str(python_dir / "pip")
                    playwright_path = str(python_dir / "playwright")
                    
                # 2. pip install -r requirements.txt
                self.root.after(0, lambda: self.append_log("Memeriksa dan memperbarui pustaka dependensi (pip install)...", "WAIT"))
                res_pip = subprocess.run([pip_path, "install", "-r", "requirements.txt"], capture_output=True, text=True, timeout=60)
                if res_pip.returncode != 0:
                    err_msg = res_pip.stderr.strip() or res_pip.stdout.strip()
                    self.root.after(0, lambda: self.append_log(f"Peringatan saat pip install: {err_msg}", "WARN"))
                else:
                    self.root.after(0, lambda: self.append_log("Instalasi dependensi Python selesai.", "OK"))
                    
                # 3. playwright install chromium
                self.root.after(0, lambda: self.append_log("Memeriksa browser Chromium Playwright...", "WAIT"))
                res_pw = subprocess.run([playwright_path, "install", "chromium"], capture_output=True, text=True, timeout=90)
                if res_pw.returncode != 0:
                    err_msg = res_pw.stderr.strip() or res_pw.stdout.strip()
                    self.root.after(0, lambda: self.append_log(f"Peringatan saat playwright install: {err_msg}", "WARN"))
                else:
                    self.root.after(0, lambda: self.append_log("Pemeriksaan browser Playwright selesai.", "OK"))
                    
                # Success
                def success_ui():
                    self.append_log("Pembaruan berhasil diterapkan! Silakan buka kembali aplikasi.", "OK")
                    for widget in self.banner_container.winfo_children():
                        widget.destroy()
                    messagebox.showinfo("Sukses", "Bot berhasil diperbarui! Silakan tutup dan buka kembali aplikasi untuk menerapkan perubahan.")
                    
                self.root.after(0, success_ui)
                
            except Exception as e:
                def fail_ui(err_err):
                    self.append_log(f"Gagal memperbarui aplikasi: {err_err}", "ERR")
                    if hasattr(self, "btn_run_update") and self.btn_run_update.winfo_exists():
                        self.btn_run_update.config(state="normal", text="📥 PERBARUI SEKARANG")
                    messagebox.showerror("Gagal Update", f"Terjadi kesalahan saat memperbarui bot:\n{err_err}")
                self.root.after(0, lambda: fail_ui(str(e)))
                
        threading.Thread(target=do_update, daemon=True).start()

    # ------------------------------------------------------------
    #  AUTO DETECTION & EXCEL BROWSE
    # ------------------------------------------------------------
    
    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Pilih File Excel Data Pasien",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            # Set values
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, file_path)
            # Trigger auto detect
            self.auto_detect_excel()
            
    def auto_detect_excel(self):
        file_val = self.get_entry_val(self.entry_excel)
        if not file_val:
            self.append_log("Nama file Excel kosong.", "WARN")
            return
            
        file_path = Path(file_val)
        if not file_path.exists():
            self.append_log(f"File Excel tidak ditemukan di path: {file_path.absolute()}", "WARN")
            return
            
        # Temporarily set the core bot's excel file path
        bot_kemkes.FILE_EXCEL = file_path
        
        # Use bot_kemkes' helper to detect last row
        self.append_log(f"Membaca file Excel: {file_path.name} ...", "INFO")
        
        # Run in a background thread to avoid locking GUI
        def run_detect():
            try:
                last_no, last_name, last_status = bot_kemkes.temukan_input_terakhir()
                
                # Update GUI safely
                if last_no is not None:
                    suggested = int(last_no) + 1
                    msg = f"Excel terdeteksi: Input terakhir No. {last_no} ({last_name}) -> Status: {last_status}. Rekomendasi Mulai: No. {suggested}."
                    
                    self.root.after(0, lambda: self.update_start_row_ui(suggested, msg))
                else:
                    msg = "Excel terdeteksi: Belum ada baris yang diproses (ditandai warna) di file ini. Rekomendasi Mulai: No. 1."
                    self.root.after(0, lambda: self.update_start_row_ui(1, msg))
            except Exception as e:
                self.root.after(0, lambda: self.append_log(f"Gagal mendeteksi input Excel terakhir: {e}", "WARN"))
                
        threading.Thread(target=run_detect, daemon=True).start()
        
    def update_start_row_ui(self, val, msg):
        self.entry_start_row.delete(0, tk.END)
        self.entry_start_row.insert(0, str(val))
        self.append_log(msg, "OK")

    # ------------------------------------------------------------
    #  LOGGING & POLLING
    # ------------------------------------------------------------
    
    def append_log(self, msg, level="INFO"):
        # Cegat sinyal internal __PATIENT_INFO__ dari bot — jangan tampilkan di log
        if msg.startswith("__PATIENT_INFO__|"):
            try:
                parts = msg.split("|")
                # Format: __PATIENT_INFO__|nomor|total|nama|nik
                nomor = int(parts[1])
                total = int(parts[2])
                nama  = parts[3]
                nik   = parts[4] if len(parts) > 4 else "-"
                self.lbl_progress.config(text=f"Memproses Pasien: {nomor} dari {total}")
                self.progress_bar.config(maximum=total, value=nomor)
                self.lbl_patient_name.config(text=nama)
                self.lbl_patient_nik.config(text=nik)
            except Exception:
                pass
            return  # jangan masuk ke log konsol

        # Format current timestamp
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        
        # Map icons
        ikon = {"INFO":"ℹ️ ","OK":"✅","WARN":"⚠️ ","ERR":"❌","WAIT":"⏳","BOT":"🤖"}.get(level,"  ")
        
        # Enable editing
        self.log_text.config(state="normal")
        
        # Insert line
        self.log_text.insert(tk.END, f"[{ts}] {ikon} {msg}\n", level)
        
        # Disable editing
        self.log_text.config(state="disabled")
        
        # Scroll to bottom
        self.log_text.yview(tk.END)
        
        # Parse progress dari isi pesan (fallback jika __PATIENT_INFO__ tidak dikirim)
        self.parse_progress_from_log(msg)

    def poll_queues(self):
        # 1. Process logs
        if self.bridge:
            while not self.bridge.log_queue.empty():
                try:
                    msg, level = self.bridge.log_queue.get_nowait()
                    self.append_log(msg, level)
                except Exception:
                    break
                    
            # 2. Process requests
            if not self.bridge.request_queue.empty():
                # Check if we are currently waiting for a request response
                if self.current_request_type is None:
                    try:
                        req_type, data = self.bridge.request_queue.get_nowait()
                        self.handle_request(req_type, data)
                    except Exception:
                        pass
                        
            # 3. Process progress / status updates
            # (We can infer progress from the logs or pass progress updates explicitly.
            #  But since the log messages contain "[nomor/total]" patterns, we can parse them,
            #  or the bot can write progress info to a state, or we can parse it directly from logs.
            #  Let's parse it from logs! It's very easy and highly dynamic.)
            
        # Re-schedule poller
        self.root.after(50, self.poll_queues)
        
    def parse_progress_from_log(self, msg):
        # Parses "[1/20]" patterns to update progress bar
        # Example: "Pasien [1/20]" or "[1/20] Akhmad selesai diproses"
        import re
        match = re.search(r'\[(\d+)/(\d+)\]', msg)
        if match:
            current = int(match.group(1))
            total = int(match.group(2))
            
            # Update progress bar
            self.progress_bar.config(maximum=total, value=current)
            self.lbl_progress.config(text=f"Memproses Pasien: {current} dari {total}")
            
        # Detect patient details
        # Example: "📋 PASIEN [1/20]" or "Nama: ... NIK: ..."
        # Wait! The bot prints "Nama          : Akhmad" and "NIK           : 12345" inside cetak_data_pasien!
        # Since these are logged, we can parse them!
        if "Nama          :" in msg:
            name = msg.split("Nama          :")[1].strip()
            self.lbl_patient_name.config(text=f"Nama Pasien: {name}")
        if "NIK           :" in msg:
            nik = msg.split("NIK           :")[1].strip()
            self.lbl_patient_nik.config(text=f"NIK: {nik}")

    # ------------------------------------------------------------
    #  BOT THREAD CONTROL
    # ------------------------------------------------------------
    
    def toggle_bot(self):
        if self.is_running:
            # Stop Request
            self.stop_bot()
        else:
            # Start Request
            self.start_bot()
            
    def start_bot(self):
        # Save current configurations to local file
        self.save_config()
        
        # 1. Gather configuration
        email = self.get_entry_val(self.entry_email)
        password = self.get_entry_val(self.entry_password)
        excel_path = self.get_entry_val(self.entry_excel)
        start_row_str = self.get_entry_val(self.entry_start_row)
        
        # Validate Excel
        if not excel_path:
            messagebox.showerror("Error", "Silakan tentukan file Excel terlebih dahulu!")
            return
        if not Path(excel_path).exists():
            messagebox.showerror("Error", f"File Excel tidak ditemukan: {excel_path}")
            return
            
        # Validate Start Row
        try:
            start_row = int(start_row_str)
        except ValueError:
            messagebox.showerror("Error", "Nomor baris mulai harus berupa angka!")
            return
            
        # 2. Update GUI elements to running state
        self.is_running = True
        self.btn_start.grid(row=12, column=0, columnspan=1, sticky="ew", pady=(10, 5))
        self.btn_start.config(text="🛑 HENTIKAN", bg=self.c_red, activebackground=self.c_red_hover)
        self.btn_start.bind("<Enter>", lambda e: self.btn_start.config(bg=self.c_red_hover))
        self.btn_start.bind("<Leave>", lambda e: self.btn_start.config(bg=self.c_red))
        
        # Create and grid SKIP PASIEN button next to it
        self.btn_skip_run = self.create_flat_button(
            self.btn_start.master, "⏭ SKIP", self.trigger_skip_run,
            bg=self.c_yellow, fg="#000000", hover_bg=self.c_yellow_hover, height=2, font=("Helvetica", 10, "bold")
        )
        self.btn_skip_run.grid(row=12, column=1, padx=(6, 0), pady=(10, 5), sticky="ew")
        
        # Disable configurations
        self.set_config_widgets_state("disabled")
        
        # Reset progress bar
        self.progress_bar.config(value=0)
        self.lbl_progress.config(text="Bot sedang memulai...")
        self.lbl_patient_name.config(text="—")
        self.lbl_patient_nik.config(text="—")
        
        # 3. Set up the GUI Bridge in bot_kemkes
        self.bridge = bot_kemkes.GUIBridge()
        self.bridge.auto_advance = self.var_auto_advance.get()
        self.bridge.default_phone = self.get_entry_val(self.entry_phone)
        
        # Activate bridge hooks in the core bot!
        bot_kemkes.GUI_BRIDGE_ACTIVE = True
        bot_kemkes.gui_bridge = self.bridge
        
        # Override the bridge's log method to also intercept progress info
        original_bridge_log = self.bridge.log
        def patched_log(msg, level="INFO"):
            original_bridge_log(msg, level)
            self.root.after(0, lambda: self.parse_progress_from_log(msg))
        self.bridge.log = patched_log
        
        # 4. Spawn the bot thread
        headless = self.var_headless.get()
        
        def bot_runner():
            self.append_log("Memulai thread Playwright Bot...", "INFO")
            try:
                # Set up the event loop for the background thread
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
                # Execute the main function of the bot with parameters
                loop.run_until_complete(
                    bot_kemkes.main(
                        username_arg=email,
                        password_arg=password,
                        mulai_baris_arg=start_row,
                        excel_file_arg=excel_path,
                        headless_arg=headless
                    )
                )
                self.append_log("Bot selesai memproses semua data pasien.", "OK")
            except Exception as e:
                self.append_log(f"Thread Bot berhenti karena error: {e}", "ERR")
            finally:
                # Reset state in GUI
                self.root.after(0, self.on_bot_finished)
                
        self.bot_thread = threading.Thread(target=bot_runner, daemon=True)
        self.bot_thread.start()
        
    def stop_bot(self):
        if not self.is_running:
            return
            
        # Set running state to False in the bridge
        if self.bridge:
            self.bridge.running = False
            # If there is a pending request, release the blocked bot thread by sending "quit"
            if self.current_request_type:
                self.send_response("q")
                
        self.append_log("Mengirim permintaan penghentian ke bot...", "WARN")
        self.lbl_progress.config(text="Bot sedang dihentikan...")
        
    def trigger_skip_run(self):
        if self.is_running and self.bridge:
            self.bridge.skip_requested = True
            self.append_log("Mengirim permintaan SKIP untuk pasien saat ini...", "WARN")
            # If there is a pending request, release it immediately
            if self.current_request_type:
                self.send_response("skip")

    def on_bot_finished(self):
        self.is_running = False
        
        # Remove the skip button from grid and destroy it
        if hasattr(self, "btn_skip_run") and self.btn_skip_run:
            try:
                self.btn_skip_run.grid_forget()
                self.btn_skip_run.destroy()
            except Exception:
                pass
            self.btn_skip_run = None
            
        # Grid btn_start back to columnspan=2
        self.btn_start.grid(row=12, column=0, columnspan=2, sticky="ew", pady=(10, 5))
        
        # Re-enable inputs
        self.set_config_widgets_state("normal")
        
        # Reset button
        self.btn_start.config(text="🚀 JALANKAN BOT", bg=self.c_green, activebackground=self.c_green_hover)
        self.btn_start.bind("<Enter>", lambda e: self.btn_start.config(bg=self.c_green_hover))
        self.btn_start.bind("<Leave>", lambda e: self.btn_start.config(bg=self.c_green))
        
        self.lbl_progress.config(text="Bot tidak aktif.")
        self.show_default_interact_message()
        
        # Deactivate bridge hooks
        bot_kemkes.GUI_BRIDGE_ACTIVE = False
        bot_kemkes.gui_bridge = None
        self.bridge = None
        self.bot_thread = None
        
        messagebox.showinfo("Informasi", "Aktivitas bot telah selesai atau dihentikan.")

    def set_config_widgets_state(self, state_val):
        self.entry_email.config(state=state_val)
        self.entry_password.config(state=state_val)
        self.entry_phone.config(state=state_val)
        self.entry_excel.config(state=state_val)
        self.entry_start_row.config(state=state_val)
        self.btn_browse.config(state=state_val)
        self.btn_detect.config(state=state_val)
        self.cb_auto_advance.config(state=state_val)
        self.cb_headless.config(state=state_val)
        self.btn_save.config(state=state_val)

if __name__ == "__main__":
    root = tk.Tk()
    app = BotGUI(root)
    
    # Elegant custom icon or window setup
    # If icon exists, load it:
    # root.iconphoto(False, tk.PhotoImage(file='icon.png'))
    
    root.mainloop()
