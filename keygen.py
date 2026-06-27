#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =============================================================
#  🔑 KEY GENERATOR - ARZACHEL BOT CKG LISENSI
#  HANYA UNTUK PENGGUNAAN ADMIN (JANGAN DIBAGIKAN KE PEMBELI!)
# =============================================================

import hashlib
import sys

# PENTING: Salt ini harus persis sama dengan yang ada di gui.py
SECRET_SALT = "arzachel_bot_ckg_salt_2026_super_secure"

def generate_activation_key(device_id):
    device_id = device_id.strip().upper()
    h = hashlib.sha256(f"{device_id}-{SECRET_SALT}".encode()).hexdigest()
    # Format key like XXXX-XXXX-XXXX-XXXX
    formatted = f"{h[12:16]}-{h[16:20]}-{h[20:24]}-{h[24:28]}".upper()
    return formatted

def print_banner():
    print("=" * 60)
    print("      🔑 ARZACHEL BOT CKG - GENERATOR KUNCI AKTIVASI      ")
    print("              (Hanya untuk Penggunaan Admin)             ")
    print("=" * 60)

def main():
    print_banner()
    try:
        while True:
            device_id = input("\n👉 Masukkan DEVICE ID pembeli (Contoh: ABCD-EFGH-IJKL): ").strip()
            if not device_id:
                print("❌ Device ID tidak boleh kosong.")
                continue
            
            # Bersihkan format
            device_id = device_id.upper()
            
            # Buat kunci aktivasi
            key = generate_activation_key(device_id)
            
            print("\n" + "-" * 50)
            print(f"📌 DEVICE ID      :  \033[1;33m{device_id}\033[0m")
            print(f"🔑 KUNCI AKTIVASI :  \033[1;32m{key}\033[0m")
            print("-" * 50)
            
            print("\nKunci di atas sudah dikaitkan khusus untuk Device ID tersebut.")
            
            tanya = input("\nIngin membuat kunci lagi? (y/n): ").strip().lower()
            if tanya != 'y':
                print("\nTerima kasih! Menutup generator.")
                break
    except KeyboardInterrupt:
        print("\n\nMenutup generator.")
        sys.exit(0)

if __name__ == "__main__":
    main()
