# Arzachel Bot CKG (ASIK Kemenkes Portal Automation)

**Arzachel Bot CKG** adalah aplikasi desktop otomasi berbasis GUI (Graphical User Interface) yang dirancang untuk melakukan registrasi data pasien secara massal (*batch registration*) pada portal resmi Kementrian Kesehatan Republik Indonesia **Sehat Indonesiaku (ASIK Kemenkes)** (`sehatindonesiaku.kemkes.go.id`) secara cepat, aman, dan efisien menggunakan pustaka **Playwright** dan **Tkinter**.

---

## 🌟 Fitur Utama

- **Antarmuka Premium Modern**: Desain GUI bernuansa *Sleek Dark Mode* yang elegan dan sangat responsif untuk kenyamanan operasional.
- **Simpan Konfigurasi Dinamis**: Fitur penyimpanan mandiri (`💾 SIMPAN KONFIGURASI`) yang menyimpan akun login (email & kata sandi), nomor WhatsApp default, dan preferensi lainnya secara lokal ke dalam berkas `gui_config.json` yang aman dan terabaikan oleh Git.
- **Deteksi Quota Habis Otomatis**: Secara instan memantau limit kuota pemeriksaan harian dari portal Kemenkes, mengklik konfirmasi kelanjutan secara otomatis, dan memberikan umpan balik langsung di log konsol.
- **Pencarian Alamat Cascade Cepat**: Menangani dropdown bertingkat portal Kemenkes (Provinsi ➔ Kabupaten ➔ Kecamatan ➔ Kelurahan) secara dinamis menggunakan pencarian teks bawaan dan *JavaScript-level clicks* untuk menghindari pemblokiran akibat backdrop/overlay.
- **Sistem Validasi Pintar**: Memiliki mekanisme polling dinamis setiap 100ms untuk menangkap popup *"Data peserta tidak valid"* atau *"Data peserta valid"* secara instan sebelum browser beralih halaman, tanpa menambah beban CPU (*zero layout-reflow*).
- **Tombol Skip Instan (Interruptible)**: Menghentikan proses pasien yang sedang aktif secara aman melalui tombol `⏭ SKIP` di GUI, mewarnai baris terkait menjadi kuning di Excel, dan segera beralih ke pasien berikutnya.
- **Pelaporan Status Warna di Excel (`ckg.xlsx`)**: 
  - 🟢 **Hijau (Success)**: Registrasi pasien berhasil selesai.
  - 🟡 **Kuning (Skipped/Invalid)**: Pasien dilewati secara manual atau data NIK tidak valid/tidak ditemukan.
  - 🔴 **Merah (Error)**: Terjadi gangguan teknis selama proses pengisian.
- **Log Konsol Rich Text**: Konsol log berwarna yang diperbarui secara *real-time* untuk membantu melacak setiap tindakan bot secara detail.

---

## 📁 Struktur Berkas Utama

```
botckg/
├── bot_kemkes.py       # Logika inti robot (Playwright, scraping, & otomasi)
├── gui.py              # Aplikasi antarmuka utama (Tkinter GUI, State & Threading)
├── gui_config.json     # File konfigurasi lokal (menyimpan kredensial secara lokal)
├── ckg.xlsx            # File Excel data pasien (Secara default diabaikan oleh Git)
├── requirements.txt    # Daftar dependensi Python
├── install.sh          # Skrip installer otomatis untuk Linux
├── install_windows.bat # Skrip installer otomatis untuk Windows
├── start_bot.sh        # Skrip sekali-klik untuk menjalankan aplikasi di Linux
└── note.md             # Catatan pengembangan dan dokumentasi teknis
```

---

## ⚙️ Instalasi & Persiapan

### 1. Klon Repositori (Clone)
Buka terminal (Linux/macOS) atau Command Prompt/PowerShell (Windows), lalu jalankan perintah berikut:

```bash
git clone https://github.com/akhmadsudaisi31-code/BOTCKG.git
cd BOTCKG
```

### 2. Pemasangan Otomatis (Instalasi)

#### 💻 Pengguna Windows:
1. Pastikan Anda telah menginstal Python 3.9 ke atas dan telah mencentang opsi **"Add Python to PATH"** saat instalasi.
2. Klik dua kali pada file `install_windows.bat` di dalam folder proyek, atau jalankan melalui Command Prompt:
   ```cmd
   install_windows.bat
   ```
3. Skrip akan secara otomatis membuat *virtual environment* (`venv`), mengunduh pustaka dependensi, dan menginstal browser Chromium yang dibutuhkan oleh Playwright.

#### 🐧 Pengguna Linux (Ubuntu/Debian):
1. Buka terminal di dalam folder proyek.
2. Berikan izin akses eksekusi pada skrip instalasi, lalu jalankan:
   ```bash
   chmod +x install.sh start_bot.sh
   ./install.sh
   ```
3. Skrip akan menginstal dependensi sistem (termasuk paket GUI Python `python3-tk` jika belum tersedia), membuat `venv`, dan memasang browser Chromium Playwright.

---

## 🚀 Cara Menjalankan Aplikasi

Setelah proses instalasi selesai dengan sukses, Anda dapat menjalankan antarmuka **Arzachel Bot CKG** menggunakan langkah berikut:

#### 💻 Pengguna Windows:
Jalankan perintah berikut di Command Prompt/PowerShell Anda:
```cmd
venv\Scripts\python gui.py
```

#### 🐧 Pengguna Linux:
Jalankan skrip sekali-klik yang telah disediakan:
```bash
./start_bot.sh
```
Atau jalankan secara manual menggunakan virtual environment:
```bash
venv/bin/python gui.py
```

---

## 📖 Panduan Penggunaan & Alur Kerja

### 1. Mempersiapkan File Excel Data Pasien (`ckg.xlsx`)
- Letakkan file Excel Anda di dalam folder utama proyek dengan nama `ckg.xlsx`.
- Format kolom minimal harus memiliki:
  - Kolom **A** berisi nomor urut (`No`) yang berfungsi sebagai penunjuk indeks baris pasien.
  - Kolom data penting lainnya seperti `NIK` (16 digit), `Nama Pasien`, `Tanggal Lahir` (format DD-MM-YYYY atau teks), jenis kelamin, dan nomor HP/WhatsApp.

### 2. Mengatur Konfigurasi Akun di Aplikasi
1. Buka aplikasi **Arzachel Bot CKG**.
2. Masukkan **Email / Username** dan **Kata Sandi** portal ASIK Kemenkes Anda pada kolom konfigurasi.
3. Tentukan **Nomor WhatsApp/HP Default** yang akan digunakan sebagai fallback jika baris pasien di Excel tidak memiliki nomor HP.
4. Klik tombol **`💾 SIMPAN KONFIGURASI`** untuk menyimpan kredensial Anda ke berkas lokal. Anda tidak perlu memasukkan kembali kredensial ini saat membuka aplikasi di lain waktu.
5. Klik **`Cari...`** untuk mengarahkan ke file Excel pasien jika lokasinya berbeda.
6. Tentukan nomor baris awal pasien yang ingin Anda proses pada kolom **"Mulai dari No. Pasien di Excel"** (atau klik **"Deteksi Lanjutan"** untuk mendeteksi baris kosong secara otomatis).

### 3. Menjalankan Proses Otomasi
1. Klik tombol hijau besar **`🚀 JALANKAN BOT`**.
2. Browser pendaftaran akan terbuka (kecuali jika Anda mencentang mode *Headless*).
3. Bot akan mengisi kredensial dan menavigasi secara otomatis.
4. ⚠️ **Verifikasi CAPTCHA Manual**: Ketika halaman login meminta CAPTCHA, antarmuka GUI akan memunculkan dialog pop-up interaktif. Isi angka CAPTCHA sesuai gambar di browser ke dalam kolom input GUI, lalu klik **Kirim/Submit**.
5. Bot akan melakukan registrasi pasien demi pasien. Anda dapat memantau progres, detail pasien aktif, dan log konsol secara langsung di layar GUI.
6. Klik **`🛑 HENTIKAN`** jika ingin mematikan bot sewaktu-waktu.

---

## 🔒 Keamanan Data & Kebijakan Git

Demi mencegah kebocoran informasi pribadi pasien (PII) dan kredensial login Kemenkes Anda ke ruang publik, berkas-berkas sensitif berikut **telah dikonfigurasi untuk diabaikan secara permanen** oleh Git melalui berkas `.gitignore`:
- `gui_config.json` (Berkas kredensial lokal Anda)
- `ckg.xlsx` dan semua file berekstensi `*.xlsx` (Data pasien)
- File cadangan excel (`*.bak`) dan file logs (`*.log`)

> [!IMPORTANT]
> Jangan pernah memaksa untuk menambahkan atau mengunggah (*commit/push*) berkas-berkas di atas ke repositori Git publik.

---

## 📜 Lisensi & Hak Cipta
Aplikasi ini dilisensikan dan dikembangkan di bawah hak cipta **Arzachel Bot CKG**. Semua hak cipta dilindungi undang-undang. Dilarang mendistribusikan ulang atau memodifikasi kode sumber ini untuk kepentingan komersial tanpa persetujuan tertulis dari pemilik hak cipta.
