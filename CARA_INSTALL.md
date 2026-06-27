# 📋 Panduan Instal Bot CKG Kemenkes
### Untuk Pengguna Awam — Ikuti langkah urut, tidak perlu mengerti pemrograman

---

## 🖥️ WINDOWS (Paling Umum)

### ✅ Yang Anda Butuhkan Sebelum Mulai
- Laptop Windows (Windows 10 / 11)
- Koneksi internet aktif
- Folder `botckg` sudah ada di laptop (dari flashdisk/Google Drive)
- Waktu ±10 menit untuk proses instalasi

---

### Langkah 1 — Instal Python *(hanya sekali, lewati jika sudah)*

> Python adalah program yang dibutuhkan agar bot bisa berjalan.

1. Buka browser, kunjungi: **https://www.python.org/downloads/**
2. Klik tombol kuning besar **"Download Python"**
3. Buka file yang diunduh
4. ⚠️ **SANGAT PENTING:** Di bagian bawah installer, **centang kotak** bertulisan:
   > **"Add Python to PATH"** atau **"Add python.exe to PATH"**
5. Klik **"Install Now"** dan tunggu hingga selesai
6. Klik **"Close"** setelah selesai

---

### Langkah 2 — Jalankan Installer Bot

1. Buka folder **`botckg`** di File Explorer
2. Klik **dua kali** file **`install_windows.bat`**
3. Jika muncul peringatan Windows (SmartScreen), klik **"More info"** lalu **"Run anyway"**
4. Ikuti petunjuk di layar jendela hitam yang muncul
5. Tunggu hingga muncul pesan **✅ INSTALASI BERHASIL!**

> ⏳ Proses ini memakan waktu 3–10 menit tergantung kecepatan internet. Jangan tutup jendela hitam!

---

### Langkah 3 — Aktivasi Bot *(Wajib, hanya 1x per laptop)*

Bot ini menggunakan sistem lisensi per perangkat. Setiap laptop perlu diaktivasi satu kali.

1. Klik **dua kali** file **`jalankan_bot.bat`**
2. Layar aktivasi akan muncul secara otomatis
3. Lihat **DEVICE ID** yang tertulis di layar (contoh: `A1B2-C3D4-E5F6`)
4. Klik tombol **"📋 SALIN"** di sebelah Device ID
5. Klik tombol hijau **"💬 HUBUNGI ADMIN VIA WHATSAPP"**
6. Kirimkan Device ID + bukti transfer ke Admin via WhatsApp
7. Admin akan membalas dengan **Kunci Aktivasi**
8. Masukkan Kunci Aktivasi ke kolom yang tersedia
9. Klik **"AKTIFKAN"** → Bot siap digunakan!

> 💳 **Biaya Lisensi:** Rp 160.000 (sekali bayar, selamanya untuk laptop ini)
> 🏦 **Transfer ke:** Bank Jago — **105295129701** a.n. Akhmad Sudaisi

---

### Langkah 4 — Siapkan Data Pasien

- Pastikan file **`ckg.xlsx`** ada di dalam folder `botckg`
- File Excel harus memiliki kolom yang benar (lihat file `template_ckg.xlsx` sebagai contoh)

---

### Langkah 5 — Jalankan Bot *(Setiap Hari)*

- Klik **dua kali** file **`jalankan_bot.bat`**
- Jendela Bot CKG akan terbuka
- Isi email dan password login Kemenkes Anda
- Klik **"Jalankan Bot"**

> 💡 Setelah terinstal dan aktif, Anda **cukup klik `jalankan_bot.bat`** setiap kali ingin menggunakan bot.

---

## ❓ Pertanyaan Umum & Solusi Masalah

---

**Q: Saat klik `install_windows.bat` langsung tertutup / tidak terjadi apa-apa?**
> A: Kemungkinan Python belum terinstal. Ulangi dari **Langkah 1** dan pastikan centang **"Add Python to PATH"**.

---

**Q: Muncul tulisan "Python tidak ditemukan" di jendela hitam?**
> A: Python belum terinstal atau lupa centang "Add Python to PATH". Instal ulang Python dari python.org, centang opsi tersebut, lalu jalankan installer lagi.

---

**Q: Proses berhenti di "Mengunduh paket program" atau "Mengunduh browser"?**
> A: Koneksi internet terputus. Pastikan internet aktif, lalu jalankan `install_windows.bat` lagi. Proses akan melanjutkan dari titik sebelumnya.

---

**Q: Muncul peringatan dari antivirus?**
> A: Nonaktifkan sementara antivirus, jalankan installer, lalu aktifkan kembali. Ini normal karena bot menggunakan browser otomatis.

---

**Q: Saya sudah bayar tapi belum dapat kunci aktivasi?**
> A: Pastikan Anda sudah mengirim **bukti transfer + Device ID** ke Admin WhatsApp. Klik tombol **"💬 HUBUNGI ADMIN VIA WHATSAPP"** di layar aktivasi untuk menghubungi langsung.

---

**Q: Apakah perlu install ulang setiap buka laptop?**
> A: **Tidak.** Installer hanya dijalankan sekali. Selanjutnya cukup klik **`jalankan_bot.bat`**.

---

**Q: Muncul error "file ckg.xlsx tidak ditemukan"?**
> A: Pastikan file Excel ada di dalam folder `botckg` dan namanya tepat **`ckg.xlsx`** (bukan ckg1.xlsx atau yang lain).

---

**Q: Browser tidak terbuka saat bot dijalankan?**
> A: Jalankan ulang **`install_windows.bat`** untuk menginstal ulang browser otomatis.

---

## 📞 Butuh Bantuan?

Hubungi Admin via WhatsApp dengan menyertakan:
- Foto / screenshot jendela hitam yang menampilkan pesan error
- Device ID Anda (terlihat di layar aktivasi bot)

---

*Bot CKG Kemenkes oleh Arzachel — sehatindonesiaku.kemkes.go.id*
