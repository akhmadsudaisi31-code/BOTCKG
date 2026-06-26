# 📋 Cara Instal Bot CKG Kemenkes di Laptop Baru

Panduan ini untuk orang awam — ikuti langkah urut, tidak perlu mengerti pemrograman.

---

## 🖥️ UNTUK LAPTOP WINDOWS

### Langkah 1 — Instal Python (hanya sekali seumur hidup)
1. Buka browser, kunjungi: **https://www.python.org/downloads/**
2. Klik tombol kuning **"Download Python 3.x.x"**
3. Jalankan file yang diunduh
4. ⚠️ **PENTING:** Centang kotak **"Add Python to PATH"** sebelum klik Install
5. Klik **Install Now** dan tunggu selesai

### Langkah 2 — Salin folder bot ke laptop
- Copy seluruh folder `botckg` ke laptop baru (via flashdisk, Google Drive, dll.)

### Langkah 3 — Jalankan Installer (hanya sekali)
1. Masuk ke folder `botckg`
2. Klik dua kali file **`install_windows.bat`**
3. Tunggu proses selesai (±3-5 menit, tergantung kecepatan internet)
4. Jika muncul jendela hitam dan ada tulisan ✅ INSTALASI BERHASIL → lanjut

### Langkah 4 — Isi Konfigurasi Login
1. Klik kanan file **`gui_config.json`** → buka dengan **Notepad**
2. Isi bagian berikut:
   ```json
   "email": "email_kemenkes_anda@gmail.com",
   "password": "kata_sandi_anda",
   ```
3. Simpan file (Ctrl+S)

### Langkah 5 — Siapkan Data Pasien
- Pastikan file **`ckg.xlsx`** ada di dalam folder `botckg`
- File Excel harus memiliki kolom: No, Nama Lengkap, NIK, Jenis Kelamin, Tanggal Lahir, dst.

### Langkah 6 — Jalankan Bot
- Klik dua kali file **`jalankan_bot.bat`**
- Jendela GUI akan terbuka → klik **Jalankan Bot**

---

## 🐧 UNTUK LAPTOP LINUX (Ubuntu/Debian)

### Langkah 1 — Salin folder bot ke laptop
```bash
# Contoh dari flashdisk
cp -r /media/namauser/FLASHDISK/botckg ~/botckg
cd ~/botckg
```

### Langkah 2 — Jalankan Installer (hanya sekali)
```bash
bash install.sh
```
Tunggu hingga muncul pesan ✅ INSTALASI BERHASIL

### Langkah 3 — Isi Konfigurasi
```bash
nano gui_config.json
```
Isi email dan password, lalu tekan **Ctrl+O** → **Enter** → **Ctrl+X**

### Langkah 4 — Jalankan Bot
```bash
bash start_bot.sh
```

---

## ❓ Pertanyaan Umum

**Q: Apakah perlu install ulang tiap buka laptop?**
A: Tidak. Installer hanya dijalankan sekali. Selanjutnya cukup klik `jalankan_bot.bat` (Windows) atau `bash start_bot.sh` (Linux).

**Q: Bot error "file ckg.xlsx tidak ditemukan"**
A: Pastikan file Excel ada di dalam folder `botckg` dan namanya tepat `ckg.xlsx`.

**Q: Browser tidak terbuka**
A: Jalankan ulang `install_windows.bat` (Windows) atau `bash install.sh` (Linux) untuk menginstal ulang Chromium.

**Q: Muncul error "Python tidak ditemukan" di Windows**
A: Instal Python dari python.org dan pastikan centang "Add Python to PATH".
