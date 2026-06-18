# Bot sehatindonesiaku.kemkes.go.id

Bot otomasi login & input data untuk portal **sehatindonesiaku.kemkes.go.id** menggunakan Playwright.

---

## 📁 Struktur File

```
bot ckg/
├── bot_kemkes.py       # Bot utama (login + input data tunggal)
├── bot_batch.py        # Input data batch dari Excel/CSV
├── debug_selector.py   # Tool debug untuk menemukan selector
├── requirements.txt    # Dependensi Python
├── data.xlsx           # (Buat sendiri) File data batch
└── screenshots/        # Screenshot otomatis tersimpan di sini
```

---

## ⚙️ Instalasi

> **Linux (Debian/Ubuntu)** — gunakan virtual environment karena pip global diblokir sistem.

```bash
# 1. Buat virtual environment
python3 -m venv venv

# 2. Install semua dependensi ke dalam venv
venv/bin/pip install -r requirements.txt

# 3. Install browser Chromium
venv/bin/playwright install chromium
```

---

## 🚀 Cara Pakai

### 1. Konfigurasi

Edit bagian atas `bot_kemkes.py`:

```python
USERNAME = "username_anda"
PASSWORD = "password_anda"
```

### 2. Temukan Selector yang Tepat

Jalankan tool debug terlebih dahulu untuk melihat struktur halaman:

```bash
venv/bin/python debug_selector.py
```

Tool ini akan mencetak semua `input` dan `button` yang ditemukan di halaman login beserta atributnya (`name`, `id`, `placeholder`).

### 3. Sesuaikan Selector

Update variabel `SEL_*` di `bot_kemkes.py` sesuai hasil debug:

```python
SEL_USERNAME  = 'input[name="username"]'
SEL_PASSWORD  = 'input[type="password"]'
SEL_BTN_LOGIN = 'button[type="submit"]'
```

### 4. Jalankan Bot

```bash
venv/bin/python bot_kemkes.py
```

**Alur eksekusi:**
1. Browser Chromium terbuka otomatis
2. Username & password terisi otomatis
3. ⏸️ **Terminal akan PAUSE** → isi CAPTCHA manual di browser
4. Tekan **ENTER** di terminal setelah CAPTCHA terisi
5. Bot mengklik tombol login
6. Bot menavigasi ke form & mengisi data

---

## 📊 Input Batch dari Excel

1. Buat file `data.xlsx` dengan kolom sesuai form (contoh: `nama`, `nik`, `tanggal`, `provinsi`)
2. Edit fungsi `input_data_batch()` di `bot_batch.py` sesuai selector form
3. Jalankan:
   ```bash
   python bot_batch.py
   ```

---

## 🐛 Debug & Troubleshooting

| Masalah | Solusi |
|---|---|
| Selector tidak cocok | Jalankan `debug_selector.py` untuk melihat semua elemen |
| Login gagal | Cek `screenshots/` folder untuk melihat kondisi halaman |
| Form tidak ditemukan | Ganti URL di `input_data()` sesuai URL form target |
| Timeout error | Naikkan nilai `TIMEOUT` (default 30000ms) |

### Mode Debug Interaktif

Tambahkan `await page.pause()` di mana saja dalam skrip untuk membuka Playwright Inspector:

```python
await page.pause()   # Buka inspector → klik elemen → salin selector
```

---

## 📸 Screenshot

Screenshot otomatis disimpan di folder `screenshots/` dengan format:
```
screenshots/YYYYMMDD_HHMMSS_nama.png
```

Tersimpan otomatis saat:
- Login berhasil
- Form terisi
- Terjadi error
