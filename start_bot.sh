#!/bin/bash
# Pindah ke direktori tempat skrip ini berada
cd "$(dirname "$0")"

# Jalankan bot dengan antarmuka GUI menggunakan venv
./venv/bin/python gui.py
