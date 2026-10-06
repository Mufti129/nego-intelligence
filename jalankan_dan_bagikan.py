#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================================================
PELUNCUR OTOMATIS: SERVER WEB PGI + CLOUDFLARE PUBLIC TUNNEL
====================================================================================================
Deskripsi:
  Menjalankan server web dashboard (Flask) sekaligus mengaktifkan Cloudflare Tunnel resmi.
  Menghasilkan link HTTPS publik resmi (trycloudflare.com) yang 100% lolos firewall kantor,
  bisa dibuka di SEMUA LAPTOP (Windows/Mac) dan SEMUA HANDPHONE (Android/iPhone).

Cara Pakai:
  python3 jalankan_dan_bagikan.py
====================================================================================================
"""

import os
import sys
import time
import re
import signal
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLOUDFLARED_BIN = os.path.join(BASE_DIR, "bin", "cloudflared")
APP_SCRIPT = os.path.join(BASE_DIR, "app_server_api.py")
PORT = 5050

if not os.path.exists(CLOUDFLARED_BIN):
    print(f"[!] Error: binary cloudflared tidak ditemukan di {CLOUDFLARED_BIN}")
    sys.exit(1)

# Pastikan permission executable
os.chmod(CLOUDFLARED_BIN, 0o755)

print("\n" + "=" * 80)
print(" 🚀 MEMULAI SERVER WEB PGI & CLOUDFLARE PUBLIC TUNNEL")
print("=" * 80)
import socket

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

proc_flask = None
if is_port_in_use(PORT):
    print(f" [✓] Server Web Dashboard sudah terdeteksi aktif di port {PORT}.")
else:
    print(f" [*] Menjalankan Server Web Dashboard di port {PORT}...")
    # 1. Jalankan Flask App Server
    proc_flask = subprocess.Popen(
        [sys.executable, APP_SCRIPT],
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True
    )
    time.sleep(1.5)

print(" [*] Mengaktifkan Cloudflare Global Tunnel (trycloudflare.com)...")

# 2. Jalankan Cloudflared Tunnel
proc_cf = subprocess.Popen(
    [CLOUDFLARED_BIN, "tunnel", "--url", f"http://localhost:{PORT}"],
    cwd=BASE_DIR,
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    bufsize=1
)

tunnel_url = None
timeout_secs = 20
start_time = time.time()

# Cari URL dari output log cloudflared
while time.time() - start_time < timeout_secs:
    line = proc_cf.stdout.readline()
    if not line:
        time.sleep(0.2)
        continue
    # Cari pola https://xxxx.trycloudflare.com
    match = re.search(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', line)
    if match:
        tunnel_url = match.group(0)
        break

if not tunnel_url:
    print("[!] Gagal mendeteksi URL tunnel dalam 20 detik.")
    print("    Silakan coba jalankan manual: ./bin/cloudflared tunnel --url http://localhost:5050")
else:
    print("\n" + "=" * 80)
    print(" 🎉 SISTEM BERHASIL AKTIF & SIAP DIBAGIKAN KE ATASAN!")
    print("=" * 80)
    print(f" 🌐 TAUTAN RESMI BISA DIBUKA DI SEMUA LAPTOP & HP:")
    print(f"    👉 {tunnel_url}")
    print("-" * 80)
    print(f" 💻 Akses di Laptop Sendiri : http://localhost:{PORT}")
    print("=" * 80)
    print("\n📋 Draf Pesan WhatsApp yang Siap Anda Salin ke Atasan:")
    print("-" * 80)
    print(f"""Selamat pagi/siang Bapak/Ibu,

Izin melaporkan, prototipe "Sistem Cerdas Simulasi & Auto-Routing Negosiasi Ruko PGI" sudah dapat dicoba langsung melalui tautan interaktif berikut:
👉 {tunnel_url}

Bisa dibuka langsung melalui browser Laptop maupun Handphone. 

Langkah Coba Singkat:
1. Ketik nama wilayah (contoh: Kota Bandung / Kab. Bekasi).
2. Masukkan harga penawaran awal ruko.
3. Klik tombol "Hitung Rekomendasi & Prediksi".

Terima kasih.""")
    print("-" * 80)
    print("\n[i] Tekan Ctrl + C di terminal ini kapan saja untuk menghentikan server.\n")

def shutdown(sig, frame):
    print("\n[*] Menghentikan server tunnel...")
    try:
        if proc_cf: proc_cf.terminate()
    except: pass
    try:
        if proc_flask: proc_flask.terminate()
    except: pass
    print("[✓] Server berhenti dengan aman.")
    sys.exit(0)

signal.signal(signal.SIGINT, shutdown)
signal.signal(signal.SIGTERM, shutdown)

# Jaga proses tetap berjalan
try:
    while True:
        time.sleep(1)
        if proc_flask is not None and proc_flask.poll() is not None:
            break
        if proc_cf is not None and proc_cf.poll() is not None:
            break
except KeyboardInterrupt:
    shutdown(None, None)
