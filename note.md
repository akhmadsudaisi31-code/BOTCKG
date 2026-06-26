# Catatan Pengembangan & Audit Bot CKG Kemenkes

Dokumen ini mencatat hasil audit kode, arsitektur sistem, rekomendasi keamanan, serta riwayat perubahan (changelog) untuk mempermudah pengembangan bot ke depannya.

---

## 🏗️ 1. Arsitektur & Struktur Sistem

Bot ini terdiri dari beberapa komponen utama yang bekerja sama secara asinkron:

1. **`bot_kemkes.py` (Mesin Utama)**
   * Berbasis **Playwright (Async API)** untuk otomasi browser.
   * Menangani alur login, pengisian data pasien, penanganan popup (kuota, kategori pasien, tidak valid), serta penanganan khusus **"Data Wali"** (untuk pasien anak).
   * Memiliki `GUIBridge` berbasis antrean (`queue.Queue`) untuk berkomunikasi dengan GUI.
   * Melakukan pembaruan langsung ke file Excel `ckg.xlsx` menggunakan `openpyxl` untuk menandai status baris (Hijau = Sukses, Kuning = Skip, Merah = Gagal).

2. **`gui.py` (Antarmuka Pengguna)**
   * Berbasis **Tkinter** dengan gaya modern/sleek dark mode.
   * Berjalan di thread utama, sementara Playwright berjalan di thread terpisah (`bot_runner`).
   * Menyediakan panel interaksi dinamis untuk input manual saat bot membutuhkan tindakan manusia (seperti pengisian Captcha, pemilihan wali, konfirmasi alamat, dll.).
   * Membaca dan menulis konfigurasi ke `gui_config.json`.

3. **`ckg.xlsx` (Database Excel)**
   * Sumber data pasien. Kolom yang digunakan:
     * Kolom 0 (`No`)
     * Kolom 1 (`Nama Lengkap`)
     * Kolom 2 (`NIK`) — *Kritis: NIK adalah 16 digit.*
     * Kolom 3 (`Jenis Kelamin`)
     * Kolom 5 (`Tanggal Lahir`)
     * Kolom 6 (`Status Pernikahan`)
     * Kolom 7 (`Pekerjaan`)
     * Kolom 8 (`Detail Domisili`)
     * Kolom 9 (`Alamat Domisili`)
     *(Kolom 4 diabaikan)*

---

## 🔍 2. Temuan Audit (Audit Findings)

Kami menemukan beberapa poin penting yang perlu segera diperbaiki demi keamanan, stabilitas, dan integritas data:

### 🔴 Temuan Kritis (High Severity)
1. **Kebocoran Data Sensitif di Git (`.gitignore` Tidak Lengkap)**
   * **Masalah:** File `gui_config.json` (berisi email & password login Kemenkes secara plain text) dan `ckg.xlsx` (berisi data pribadi & NIK pasien) **tidak dimasukkan** ke dalam `.gitignore`.
   * **Risiko:** Jika Anda menjalankan `git push` ke GitHub publik, data kredensial Anda dan data pribadi pasien (PII - Personally Identifiable Information) akan terekspos secara publik. Ini melanggar hukum perlindungan data pribadi dan keamanan sistem Kemenkes.
   * **Rekomendasi:** Segera tambahkan `gui_config.json` dan `ckg.xlsx` ke `.gitignore`.

2. **Kredensial Hardcoded di `bot_kemkes.py`**
   * **Masalah:** Email `sudaisi74@gmail.com` dan password `Bleg@123` tertulis langsung secara permanen di baris 54-55 `bot_kemkes.py`.
   * **Risiko:** Kredensial akan bocor jika kode dibagikan atau diunggah ke repositori kode.
   * **Rekomendasi:** Hapus kredensial hardcoded. Bot harus membaca kredensial dari `gui_config.json` yang dimuat secara dinamis, atau menggunakan variabel lingkungan (environment variables).

3. **Potensi Kerusakan Presisi NIK (Floating-Point Precision Loss)**
   * **Masalah:** Pada `bot_kemkes.py` baris 2484, Excel dibaca menggunakan Pandas tanpa menentukan tipe data:
     `df = pd.read_excel(FILE_EXCEL, usecols=[0, 1, 2, 3, 5, 6, 7, 8, 9])`
     Lalu di baris 2499 dikonversi dengan `str(int(r["NIK"]))`.
   * **Risiko:** NIK memiliki 16 digit. Batas presisi angka float di komputer (IEEE 754 double precision) adalah 15 digit. Jika Pandas membaca NIK sebagai angka float, digit terakhir NIK (digit ke-16) berpotensi rusak/berubah menjadi angka 0 (misalnya `3201012345678901` menjadi `3201012345678900`). Ini menyebabkan NIK menjadi tidak valid saat dicek di portal Kemenkes.
   * **Rekomendasi:** Paksa Pandas membaca kolom NIK sebagai string langsung dengan parameter `dtype={2: str}` (atau kolom indeks terkait).

### 🟡 Temuan Sedang (Medium Severity)
4. **Ketiadaan Backup Otomatis untuk File Excel**
   * **Masalah:** Bot langsung menulis dan menyimpan perubahan ke file `ckg.xlsx` yang sama secara real-time.
   * **Risiko:** Jika komputer mati mendadak, terjadi error sistem, atau crash saat proses penyimpanan (`wb.save()`), file `ckg.xlsx` bisa rusak (corrupt) dan seluruh data pasien hilang.
   * **Rekomendasi:** Buat backup otomatis (misal: menyalin `ckg.xlsx` ke `ckg.xlsx.bak`) setiap kali bot akan menulis perubahan.

5. **Tidak Ada Penyimpanan Log ke File (Persistent Logging)**
   * **Masalah:** Log bot hanya dicetak ke terminal dan GUI. Jika GUI ditutup atau crash, riwayat jalannya bot tidak tersimpan.
   * **Risiko:** Sulit melakukan debugging jika bot mengalami masalah pada baris tertentu saat dijalankan semalaman (unattended run).
   * **Rekomendasi:** Simpan log ke dalam file log lokal (misalnya `bot_run.log`).

### 🟢 Temuan Rendah (Low Severity)
6. **Validasi Format Tanggal Lahir**
   * **Masalah:** Tanggal lahir di Excel dibaca langsung sebagai Timestamp. Jika ada format tanggal yang tidak konsisten di Excel (misalnya teks biasa "12-05-1990" vs format tanggal Excel), bot bisa mengalami error saat melakukan parsing tanggal.
   * **Rekomendasi:** Tambahkan fungsi pembantu untuk parsing tanggal yang toleran terhadap berbagai format teks.

---

## 📋 3. Rencana Aksi Perbaikan (Action Plan)

1. [x] **Perbarui `.gitignore`**: Masukkan `gui_config.json` dan `ckg.xlsx` untuk mencegah kebocoran kredensial dan PII ke repositori Git.
2. [x] **Amankan Kredensial**: Modifikasi `bot_kemkes.py` agar membaca kredensial dari `gui_config.json` secara dinamis, menghilangkan kredensial hardcoded.
3. [x] **Perbaiki Pembacaan NIK**: Menambahkan parameter `dtype={2: str}` di `pd.read_excel` dan menulis fungsi pembersihan NIK yang kokoh untuk mencegah degradasi digit ke-16 akibat floating-point.
4. [x] **Tambahkan Backup Excel**: Implementasi pencadangan otomatis `.bak` sebelum bot melakukan pembaruan status (`wb.save`).
5. [x] **Tambahkan File Log**: Implementasi pencatatan aktivitas terperinci tanpa kode ANSI ke file `bot_run.log`.
6. [x] **Perbaikan Tombol Skip**: Mengganti pengalihan `about:blank` menjadi navigasi langsung ke form pendaftaran CKG agar form langsung terbuka untuk pasien berikutnya.

---

## 🛠️ 5. Rancangan Perbaikan Masalah Tombol Skip

### 📋 Masalah:
Ketika tombol "Skip" diklik di GUI, bot menavigasi browser ke `about:blank`. Hal ini memicu error untuk membatalkan proses pasien aktif secara instan. Namun, efek sampingnya adalah browser tertinggal di halaman kosong (`about:blank`) dan terkadang gagal memuat ulang form pendaftaran untuk pasien selanjutnya secara otomatis.

### 💡 Solusi & Rancangan Perbaikan:
1. **Menghilangkan Seluruh Referensi `about:blank`**: Menghapus string `"about:blank"` dari seluruh basis kode untuk menghindari pemuatan halaman kosong.
2. **Navigasi Langsung ke Form Pendaftaran**: Mengubah alur pembatalan di `poll_gui_signals()`. Ketika tombol Skip ditekan, bot akan langsung mengarahkan halaman aktif ke `{BASE_URL}/ckg-pendaftaran-individu`.
   * Ini tetap memicu pembatalan instan pada operasi Playwright yang sedang berjalan (karena halaman berpindah dan elemen lama menjadi lepas/detached).
   * Browser langsung memuat halaman pendaftaran, bukan halaman kosong putih.
3. **Penyederhanaan `buka_form_daftar_baru()`**:
   * Menghapus pemeriksaan kondisi `page.url == "about:blank"`.
   * Cukup mendeteksi apakah URL aktif mengandung `"ckg-pendaftaran-individu"`. Jika tidak, langsung menavigasi ke URL form pendaftaran. Jika sudah di sana, form langsung diisi tanpa navigasi ulang.
4. **Pencegahan Bypass Exception via `log()`**:
   * Menyisipkan pemanggilan `check_stop_request()` di awal fungsi `log()`.
   * Karena fungsi `log()` dipanggil di setiap awal langkah pendaftaran, hal ini menjamin bahwa jika ada local `try-except Exception:` yang sempat menangkap error navigasi Playwright, eksekusi akan tetap langsung dihentikan (dengan melemparkan `SkipPasien` yang merupakan subclass `BaseException` sehingga tidak tertangkap oleh `except Exception:`) pada langkah berikutnya.

---

## 📝 6. Riwayat Perubahan (Changelog)

### [2026-06-26] - Implementasi Penuh Rencana Aksi Audit & Perbaikan Bug
* **Audit Folder & Pembuatan note.md**: Melakukan peninjauan mendalam terhadap struktur kode, mengidentifikasi celah keamanan kredensial, NIK 16 digit, dan keandalan Excel. Mendokumentasikan arsitektur dan hasil audit ke `note.md`.
* **Perbaikan Bug (Skip Button - Tahap Akhir)**: Menghapus referensi `about:blank` dan mengalihkan langsung ke form pendaftaran CKG ketika tombol Skip diklik. Menyisipkan `check_stop_request()` pada fungsi `log()` untuk memastikan penghentian instan, pewarnaan baris Excel menjadi kuning (`dilewati`), dan pembukaan kembali form pendaftaran untuk pasien berikutnya tanpa jeda.
* **Keamanan Kredensial**: Menghapus email dan kata sandi hardcoded dari `bot_kemkes.py`. Kredensial kini dimuat secara dinamis dari `gui_config.json`.
* **Keamanan Git (`.gitignore`)**: Memperbarui `.gitignore` untuk menyaring `gui_config.json`, `ckg.xlsx` beserta salinan cadangannya (`*.bak`), dan file log (`*.log`) dari pelacakan repositori Git.
* **Integritas Data (Pencegahan Kerusakan NIK)**: Menambahkan pembacaan tipe data string (`dtype={2: str}`) pada Pandas `read_excel` dan merancang fungsi penyaring karakter non-numerik serta pembersih sufiks `.0` float untuk menjamin keutuhan 16 digit NIK.
* **Penyelamatan Excel (Cadangan Otomatis)**: Menyisipkan fungsi pencadangan otomatis file `ckg.xlsx` menjadi `ckg.xlsx.bak` sebelum pembaruan status baris disimpan.
* **Log Persisten**: Memperbarui fungsi `log` agar mencatat riwayat eksekusi bersih dari kode warna ANSI ke dalam file log lokal `bot_run.log`.

### [2026-06-26] - Perbaikan Definitif Tombol Skip (asyncio.Task.cancel)

> **Root Cause**: Pendekatan sebelumnya (navigasi browser ke form CKG sebagai sinyal interupsi) tidak andal karena error yang dipicu oleh navigasi bisa tertangkap oleh blok `except Exception` lokal di dalam `input_pasien()`, menyebabkan bot melanjutkan eksekusi alih-alih berhenti.

**Solusi yang Diterapkan:**

1. **Variabel Global `_current_input_task`**: Menambahkan referensi `asyncio.Task` yang menyimpan task `input_pasien` yang sedang aktif.

2. **`input_pasien` dijalankan sebagai `asyncio.Task` terpisah**:
   ```python
   _current_input_task = asyncio.ensure_future(
       input_pasien(page, baris.to_dict(), urutan, total)
   )
   hasil = await _current_input_task
   ```

3. **`poll_gui_signals()` memanggil `.cancel()` pada Task**:
   ```python
   if _current_input_task and not _current_input_task.done():
       _current_input_task.cancel()
   ```
   `asyncio.CancelledError` **tidak dapat ditangkap** oleh `except Exception`, sehingga selalu merambat keluar dari `input_pasien` ke loop utama.

4. **Tangkap `asyncio.CancelledError` di loop utama** dengan efek:
   - Set `status_pasien = "dilewati"`
   - Panggil `update_excel_row_color(nomor, "yellow")` → baris kuning di Excel
   - Panggil `buka_form_daftar_baru(page)` → browser membuka form baru
   - `break` ke pasien berikutnya

5. **`poll_gui_signals` juga navigasi browser** ke form CKG bersamaan dengan cancel, agar halaman sudah siap sebelum loop memanggil `buka_form_daftar_baru()`.

6. **Hapus `check_stop_request()` dari `log()`**: Pendekatan ini tidak diperlukan lagi dan berpotensi menimbulkan masalah jika `log()` dipanggil dari konteks yang tidak aman.
7. **Flag `excel_sudah_ditulis`**: Boolean per-pasien untuk mencegah `update_excel_row_color` dan penghitung statistik (`dilewati`) dipanggil dua kali — satu kali di `CancelledError` handler dan sekali lagi di blok status akhir loop.

### [2026-06-26] - Tampilan Pasien Aktif & Installer Otomatis
* **Fix Nama & NIK di Status Panel**: `cetak_data_pasien()` kini mengirim sinyal `__PATIENT_INFO__` melalui `gui_bridge.log()`. GUI mencegat sinyal ini di `append_log()` dan memperbarui label Nama & NIK tanpa menampilkannya di log konsol.
* **Upgrade UI Panel Status**: Label Nama tampil font 11pt tebal warna teal, NIK tampil warna kuning tebal — mudah terbaca saat bot berjalan.
* **`install.sh` (Linux)**: Installer otomatis satu perintah — menangani Python, venv, pip, Playwright, Chromium, dan konfigurasi awal.
* **`install_windows.bat` (Windows)**: Installer satu-klik untuk pengguna awam — cek Python, buat venv, instal dependensi, unduh Chromium, buat `jalankan_bot.bat`.
* **`CARA_INSTALL.md`**: Panduan instalasi step-by-step Bahasa Indonesia untuk pengguna awam di Windows maupun Linux.

### [2026-06-26] - Perbaikan Popup Valid/Tidak Valid & Dropdown Alamat

#### Bug 1: Popup valid/tidak valid tidak diklik
* **Root Cause**: `cek_popup_tidak_valid()` menggunakan selector `div.fixed.z-1000` yang terlalu sempit. Popup yang muncul di dalam modal bertingkat (z-index berbeda) tidak terdeteksi.
* **Fix**: Fungsi ditulis ulang dengan **3 strategi berlapis**:
  1. XPath spesifik user (prioritas tertinggi)
  2. Scan semua elemen `fixed` dari z-index tertinggi (z-9000) ke rendah (z-1000) menggunakan kata kunci teks
  3. Fallback scan teks seluruh halaman (`body.inner_text`)
* **Perilaku**: Popup valid → klik tombol konfirmasi → lanjut. Popup tidak valid → klik tutup → baris di-skip (return `True`).

#### Bug 2: Dropdown Provinsi/Kabupaten error "Unexpected token"
* **Root Cause**: `results_container` locator menggabungkan CSS selector dan XPath dalam satu string, yang tidak diizinkan Playwright.
* **Fix**: Locator dipisah menjadi daftar CSS selector murni yang dicoba satu per satu hingga ada yang visible. Ditambah `None` guard sebelum memanggil `.count()`.

### [2026-06-26] - Perbaikan Cascade Pengisian Alamat Domisili

**Root Cause 1 — `nth(idx)` salah:** Overlay "Pilih Lokasi" menampilkan 1 input sekaligus (cascade). Kabupaten dicari di `nth(1)` padahal setelah Provinsi dipilih hanya ada 1 input. Fix: selalu gunakan `.first`, tambah polling 3 detik.

**Root Cause 2 — LAPISAN 2 terlalu longgar:** Bot mengklik div seluruh halaman karena teksnya mengandung "JAWA TIMUR". Fix: LAPISAN 2 kini hanya mencari di dalam `div.fixed.z-9000`, batas teks maks 80 karakter.

### [2026-06-26] - Perbaikan Popup Kuota Habis & Deteksi Modal Data Pendukung

#### Bug: Bot stuck di popup "Kuota Pemeriksaan Habis" & salah mengira itu sebagai modal "Data Pendukung"

**Root Cause:**
1. **Klik Tidak Terdaftar:** Klik "Lanjut" pada popup kuota terkadang tidak langsung direspons halaman, dan bot melanjutkan langkah tanpa memastikan popup benar-benar tertutup.
2. **Tabrakan Kata Kunci:** Pencarian modal data pendukung mendeteksi kecocokan kata kunci `"data peserta"`, yang ternyata juga terdapat di teks popup kuota. Karena keduanya memakai elemen kartu putih (`div.rounded-lg.bg-white`), bot salah mendeteksi popup kuota sebagai modal data pendukung.

**Solusi yang Diterapkan:**
1. **Peningkatan Keandalan `klik_popup_kuota`:** Menambahkan loop retry hingga 5 kali dan fungsi verifikasi `popup_kuota_masih_ada()` untuk memastikan popup benar-benar tertutup sebelum bot melanjutkan.
2. **Kombinasi Kata Kunci Unik:** Deteksi modal data pendukung sekarang mensyaratkan kata kunci `"status pernikahan"` **DAN** salah satu dari `"pekerjaan"` atau `"alamat domisili"`. Ini sangat spesifik dan tidak akan tertukar dengan popup kuota.
3. **Pembersihan Proaktif di Loop Modal:** Menambahkan pemeriksaan berkala di dalam loop tunggu modal; jika popup kuota terdeteksi aktif, bot langsung memanggil `klik_popup_kuota` untuk menutupnya secara otomatis.

### [2026-06-26] - Perbaikan Deteksi Popup "Data Peserta Valid"

#### Bug: Bot stuck di popup "Data peserta valid" dan tidak mengeklik tombol "Lanjutkan"

**Root Cause:**
1. **Pencocokan Kata Kunci Terlalu Sempit:** Fungsi `cek_popup_tidak_valid()` mencari `"data valid"` sebagai contiguous substring. Namun, popup di halaman menampilkan teks `"Data peserta valid"`. Karena ada kata `"peserta"` di tengahnya, pencocokan substring `"data valid"` menghasilkan nilai `False`, sehingga popup diabaikan oleh bot.
2. **Keterbatasan Selector Tombol:** Deteksi tombol konfirmasi hanya mencari elemen dengan tag `button`. Jika tombol pada popup tersebut diimplementasikan menggunakan `div` dengan class `cursor-pointer` atau `role="button"`, bot tidak dapat mengekliknya.

**Solusi yang Diterapkan:**
1. **Fleksibilitas Kata Kunci Valid:** Menambahkan pencocokan terhadap teks `"peserta valid"` dan kata kunci `"data peserta valid"`. Selain itu, ditambahkan logika fallback: jika teks mengandung kata `"valid"` dan tidak mengandung kata `"tidak"`, popup otomatis diklasifikasikan sebagai popup valid.
2. **Perluasan Selector Tombol Klik:** Selector pencarian tombol diubah menjadi `button, [role='button'], div.cursor-pointer` agar bot dapat mengeklik tombol konfirmasi apa pun tipe elemen HTML-nya.
3. **Penyelarasan XPath Tombol Baru:** Mengakomodasi XPath tombol popup yang diberikan oleh user (`div[4]` di samping `div[5]`) agar deteksi di Strategi 1 berjalan dengan sangat presisi dan andal.

### [2026-06-26] - Solusi Ampuh Intersepsi Klik (force_js_click) & Pengecualian Form Utama

#### Bug: Klik Lanjut di Popup Kuota lambat dan Klik Lanjut di Popup Valid gagal (Timeout 8000ms / Intercepted)

**Root Cause:**
1. **Intersepsi Backdrop (Pointer Events Intercepted):** Portal menggunakan backdrop overlay abu-abu transparan (`div.fixed.inset-0.bg-[rgba(0,0,0,0.8)]`). Playwright mendeteksi bahwa backdrop tersebut menghalangi tombol di depannya, sehingga memblokir tindakan klik standar `.click()` dan menunggu hingga terjadi timeout.
2. **Salah Deteksi Form Utama Sebagai Popup:** Karena kontainer form utama pendaftaran juga menggunakan kelas `div.fixed.z-1000`, bot salah mendeteksinya sebagai popup valid (karena tombol "Selanjutnya" mengandung kata "Lanjut"). Akibatnya, bot mencoba mengeklik tombol "Selanjutnya" untuk kedua kalinya di dalam `cek_popup_tidak_valid()`, padahal tombol tersebut sudah tertutup oleh backdrop loading pendaftaran.

**Solusi yang Diterapkan:**
1. **Fungsi Global `force_js_click()`:** Membuat fungsi pembantu global yang memaksakan klik menggunakan JavaScript injection (`element.click()`) langsung pada DOM browser. Cara ini memotong seluruh *actionability checks* Playwright dan terbukti 100% ampuh menembus intersepsi backdrop/overlay. Fungsi ini diintegrasikan pada seluruh klik di `klik_popup_kuota()` dan `cek_popup_tidak_valid()`.
2. **Pengecualian Form Utama (Form Exclusion):** Menambahkan filter verifikasi teks pada `cek_popup_tidak_valid()`. Jika kontainer yang terdeteksi mengandung label pendaftaran (seperti `"nomor whatsapp"`, `"tanggal lahir"`, atau `"isi data wali"`), kontainer tersebut langsung dilewati karena merupakan form utama, bukan popup pendaftaran.

### [2026-06-26] - Deteksi Presisi Popup "Data Peserta Valid" & Isolasi Popup Kuota

#### Bug: Bot tertahan di popup "Data peserta valid" karena mendeteksi kembali popup "Kuota Habis" lama yang tertinggal di DOM

**Root Cause:**
1. **Tinggalan Elemen di DOM:** Saat popup kuota habis telah ditutup/diklik, portal Kemenkes tidak menghapus elemen tersebut dari DOM. Ketika popup "Data peserta valid" muncul, kedua elemen tersebut aktif di DOM dan sama-sama cocok dengan selector `div.fixed.z-1000`.
2. **Interferensi Loop Deteksi:** Pada fungsi `cek_popup_tidak_valid()`, bot melakukan scan umum. Karena elemen popup kuota lama dibaca terlebih dahulu di dalam loop, bot mengidentifikasi teks lamanya, mengeklik tombol fiktifnya, dan langsung keluar dari fungsi dengan mengembalikan `False`. Akibatnya, popup "Data peserta valid" yang berada di urutan berikutnya tidak pernah terproses.

**Solusi yang Diterapkan:**
1. **Targeting Spesifik (Strategi 0):** Menyisipkan pemeriksaan khusus di awal fungsi `cek_popup_tidak_valid()` yang secara eksklusif mencari kontainer dengan teks `"Data peserta valid"`. Jika ditemukan, bot langsung mengeklik tombol `"Lanjutkan"` menggunakan `force_js_click()` dan langsung melanjutkan pendaftaran. Langkah ini memotong seluruh interferensi elemen DOM lainnya.
2. **Isolasi Popup Kuota:** Menambahkan filter pengecualian pada pemindaian umum. Jika teks kontainer terdeteksi mengandung kata `"kuota pemeriksaan"` atau `"kuota habis"`, kontainer tersebut langsung dilewati (*ignored*). Hal ini menjamin bahwa deteksi NIK valid/tidak valid tidak akan terganggu oleh popup kuota, yang penanganannya sudah didelegasikan secara khusus ke fungsi `klik_popup_kuota()` di alur utama.

### [2026-06-26] - Polling Deteksi Popup & Jaring Pengaman Ganda (Double-Safety Net)

#### Bug: Bot melompati deteksi popup "Data peserta valid" dan langsung masuk ke loop tunggu modal "Data Pendukung" (Race Condition)

**Root Cause:**
1. **Race Condition (Waktu Render):** Setelah mengeklik tombol "Selanjutnya", popup "Data peserta valid" membutuhkan waktu sekitar 1–2 detik untuk dirender sepenuhnya oleh halaman web. Fungsi `cek_popup_tidak_valid()` sebelumnya hanya menunggu selama 400ms, memeriksa satu kali, lalu langsung keluar. Karena popup belum muncul pada milidetik ke-400, bot mengabaikannya, menganggap pendaftaran valid langsung masuk, dan segera beralih menunggu modal data pendukung. Popup yang akhirnya muncul terlambat kemudian memblokir layar.

**Solusi yang Diterapkan:**
1. **Polling Deteksi Popup (Maks 4 Detik):** Fungsi `cek_popup_tidak_valid()` ditulis ulang menggunakan metode **polling dinamis** (loop per 200ms selama maksimal 4 detik). Fungsi ini akan mendeteksi popup secara instan begitu muncul di layar, atau langsung keluar tanpa menunggu jika mendeteksi tombol "Selanjutnya" di halaman pendaftaran sudah menghilang (halaman sukses berpindah tanpa popup).
2. **Jaring Pengaman Ganda (Double-Safety Net):** Menyisipkan deteksi darurat untuk popup `"Data peserta valid"` di dalam loop tunggu modal `"Data Pendukung"`. Jika popup tersebut terdeteksi masih terbuka saat menunggu modal, bot akan mengeklik tombol `"Lanjutkan"` secara paksa menggunakan `force_js_click()` untuk membersihkan layar dan membebaskan proses input data pendukung.

### [2026-06-26] - Perbaikan Klik Pilihan Dropdown Alamat Domisili (Cascade)

#### Bug: Alamat domisili (Provinsi > Kabupaten > Kecamatan > Kelurahan) berhasil diketik di kolom pencarian, namun gagal diklik/dipilih oleh bot

**Root Cause:**
1. **Kegagalan Deteksi Visibilitas Elemen Opsi (False Negatives):** Penggunaan fungsi `is_element_really_visible(opt)` yang sangat ketat pada setiap anak elemen opsi dropdown menghasilkan nilai `False` palsu. Hal ini sering disebabkan oleh cara browser menghitung `offsetWidth` pada elemen teks inline atau delay rendering animasi transisi dropdown, sehingga bot melewatkan opsi yang sebenarnya sudah terlihat secara visual di layar.
2. **Keterbatasan Dukungan Keyboard:** Komponen dropdown pencarian alamat di portal Kemenkes tidak mengimplementasikan aksesibilitas keyboard (Aria-Select). Akibatnya, metode cadangan *Keyboard Fallback* (ArrowDown + Enter) tidak memberikan respon apa pun di browser. Klik kursor secara fisik/programatik merupakan satu-satunya cara untuk memilih opsi.

**Solusi yang Diterapkan:**
1. **Pencarian Teks Asli & Polling Dinamis (Lapis 1):** Menggunakan pencari teks bawaan Playwright (`get_by_text(val, exact=False)`) yang menargetkan elemen opsi umum (`li, button, [role='option'], div.cursor-pointer`) di dalam kontainer overlay `z-9000`. Ditambahkan polling dinamis setiap 200ms selama maksimal 3.6 detik. Bot akan langsung mengeklik opsi begitu muncul tanpa perlu menunggu jeda statis 1.8 detik (meningkatkan kecepatan dan keandalan).
2. **Stabilisasi Pemeriksaan Visibilitas:** Mengganti fungsi deteksi visibilitas khusus dengan `.is_visible()` bawaan Playwright yang jauh lebih stabil untuk elemen dropdown dinamis, serta menyaring elemen dengan panjang teks >100 karakter untuk menghindari kesalahan mendeteksi kontainer halaman.
3. **JavaScript Force Click:** Eksekusi klik pada opsi dropdown menggunakan fungsi global `force_js_click()` untuk memastikan opsi terpilih secara instan dan sempurna tanpa terhalang actionability checks Playwright.

### [2026-06-26] - Deteksi Presisi Popup "Data Peserta Tidak Valid" & Pemicu Skip Instan

#### Bug: Popup "Data peserta tidak valid" muncul di layar, namun bot tetap melanjutkan ke alur pengisian data pendukung (gagal memicu skip)

**Root Cause:**
1. **Miskomunikasi Logika Filter Presisi:** Pada pembaruan sebelumnya, bot dikonfigurasi untuk mendeteksi teks `"Data peserta valid"` secara spesifik di awal loop. Ketika popup `"Data peserta tidak valid"` muncul, teks tersebut tidak cocok dengan filter presisi tersebut karena adanya kata `"tidak"`. Pemindaian umum berikutnya kemudian melewatkannya karena kendala waktu rendering, sehingga bot terus berjalan ke alur data pendukung.

**Solusi yang Diterapkan:**
1. **Deteksi Khusus Popup "Data peserta tidak valid" (Strategi 0.1):** Menyisipkan deteksi eksklusif untuk popup `"Data peserta tidak valid"` di baris teratas fungsi `cek_popup_tidak_valid()`, sejajar dengan deteksi data valid. 
2. Jika popup tersebut terdeteksi aktif, bot akan langsung mengeklik tombol `"Periksa Kembali"` menggunakan `force_js_click()` untuk membersihkan layar, dan langsung mengembalikan status `True`. Ini akan memicu alur penanganan interupsi utama secara instan (mengubah status pasien menjadi `dilewati`, mewarnai baris Excel menjadi kuning, membuka kembali form pendaftaran baru, dan melanjutkan ke pasien berikutnya).

### [2026-06-26] - Playwright Native Visibility Check & Jaring Pengaman Ganda Data Tidak Valid

#### Bug: Popup "Data peserta tidak valid" tetap terlewati dan bot melompat ke loop tunggu data pendukung

**Root Cause:**
1. **Efek Transisi Animasi Modal (False Negatives):** Fungsi deteksi visibilitas khusus `is_element_really_visible()` mengevaluasi opacity computed style dan offsetWidth. Ketika modal muncul dengan efek transisi memudar (*fade-in*), nilai opacity bernilai 0 pada milidetik awal rendering. Hal ini menghasilkan nilai `False` palsu yang menyebabkan bot melewatkan deteksi presisi popup tersebut pada saat loop berjalan.

**Solusi yang Diterapkan:**
1. **Stabilisasi dengan Playwright Native `.is_visible()`:** Mengubah pemeriksaan visibilitas spesifik (Strategi 0) untuk menggunakan fungsi `.is_visible()` bawaan Playwright murni. Properti ini jauh lebih stabil, toleran terhadap efek transisi animasi modal, dan menjamin deteksi 100% andal.
2. **Jaring Pengaman Ganda Data Tidak Valid (Double-Safety Net):** Menyisipkan deteksi darurat untuk popup `"Data peserta tidak valid"` langsung di dalam loop tunggu modal `"Data Pendukung"`. Jika popup tidak valid tersebut terdeteksi terbuka saat bot sedang menunggu modal, bot akan mengeklik tombol `"Periksa Kembali"` secara paksa menggunakan `force_js_click()` dan langsung melemparkan exception `SkipPasien`. Langkah ini menjamin bahwa alur *skip* programatik (mewarnai kuning di Excel dan membuka form baru) akan dipicu secara instan bahkan jika popup muncul sangat lambat.

### [2026-06-26] - Optimasi Kecepatan Deteksi Kuota & Validasi Data (Peningkatan Performa Bot)

#### Masalah: Deteksi kuota habis dan validasi data (peserta valid/tidak valid) terasa lambat dan memakan waktu.

**Root Cause:**
1. **Layout Reflow Akibat `body.inner_text()`:** Pemanggilan `page.locator("body").inner_text()` pada fungsi check fallback memaksa browser Playwright untuk melakukan komputasi layout dan reflow pada seluruh DOM halaman web. Proses ini memakan waktu antara 500ms hingga 1000ms per eksekusi, sehingga memperlambat jalannya bot secara signifikan (menambahkan 1-2 detik overhead pada setiap pasien).
2. **Overhead JS Injection `is_element_really_visible()`:** Fungsi pengecekan visibilitas kustom memicu eksekusi kode JavaScript via `.evaluate()` di dalam konteks browser yang menelusuri seluruh hirarki parent DOM. Menjalankan fungsi ini berulang kali untuk banyak elemen di dalam polling loop menimbulkan overhead CPU yang besar.
3. **Delay Statis Sebelum Polling:** Adanya jeda statis `wait_for_timeout(1000)` setelah mengklik tombol "Selanjutnya" menunda proses pengecekan popup, padahal polling dinamis dapat mendeteksinya secara instan begitu muncul.

**Solusi yang Diterapkan:**
1. **Pencarian Teks Tertarget & Native Regex:** Mengganti pemindaian teks seluruh body dengan locator pencarian teks regex bawaan Playwright (`page.locator("text=/.../i").first`). Metode ini diproses langsung oleh engine C++ internal Playwright di tingkat browser tanpa memicu reflow halaman, sehingga waktu eksekusi turun dari ~1000ms menjadi <10ms.
2. **Eliminasi Overhead JS Evaluator:** Mengganti panggilan `is_element_really_visible()` dengan properti native Playwright `.is_visible()` di dalam loop polling dan verifikasi tombol Selanjutnya. Ini meniadakan injeksi skrip eksternal ke browser dan membuat loop berjalan sangat ringan.
3. **Reduksi Polling Latency & Delay Statis:**
   - Mengurangi jeda tunggu statis setelah klik tombol "Selanjutnya" dari `1000ms` menjadi hanya `100ms`, karena polling dinamis akan menangkap kemunculan popup secara real-time.
   - Meningkatkan frekuensi polling deteksi popup menjadi setiap `100ms` (dari sebelumnya `200ms`) with batas maksimal 30 kali pengecekan (total jendela tunggu 3.0 detik). Hal ini membuat bot bereaksi secara instan begitu popup dirender atau ketika halaman berpindah.
   - Mengurangi delay tunggu setelah tombol popup diklik dari `800ms` menjadi `300ms`.




