# Pusat Gadai Indonesia (PGI) — Negotiation Intelligence & Expansion Decision Support System

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Sistem analitik terpadu korporasi berbasis kecerdasan data (Data Intelligence) untuk evaluasi kinerja negosiasi ruko cabang (UPC), dekomposisi siklus lead time, pemetaan geospasial kemudahan negosiasi wilayah, audit risiko vendor renovasi, serta **Smart Auto-Routing Engine** berbasis Multi-Criteria Decision Analysis (MCDA) untuk penugasan negosiator ruko baru secara objektif, proporsional, dan terukur.

---

## Modul Utama Sistem

### 1. Overview & KPI Eksekutif
- **Metrik Kunci:** Pelacakan 1.039 cabang database, total efisiensi biaya riil (Rp 2,13 Miliar), rata-rata yield diskon (15.60%), dan SLA durasi negosiasi (11.8 hari vs target 17 hari).
- **Komparasi Kinerja Tim:** Analisis kinerja **Bonita** (High-Value Specialist: 142 deal, Rp 1,40 M saving, 18.25% diskon) vs **Mirza** (Speed Specialist: 132 deal, Rp 728,4 Juta saving, 14.05% diskon) vs **Tim Pendamping**.
- **Visualisasi:** Boxplot sebaran efisiensi diskon dan scatter plot harga penawaran awal vs harga deal final.

### 2. Smart Auto-Routing & Evaluator AI
- **Multi-Criteria Decision Analysis (MCDA):** Mengintegrasikan 4 pilar utama:
  1. *Workload Capacity & Fair Distribution* (Kapasitas aktif tim ≤ 15 proyek)
  2. *Regional Historical Track Record* (Stronghold historis wilayah)
  3. *Ticket Size / Value Affinity* (Nilai penawaran ruko ≥ Rp 50 Juta)
  4. *SLA Speed Optimization* (Target SLA ≤ 17 hari)
- **Output:** Rekomendasi negosiator resmi, rasional keputusan otomatis, estimasi nominal penghematan (Rp), estimasi diskon (%), estimasi durasi (hari), dan Radar Chart komparasi skor.

### 3. Geospatial & Analisis Regional
- **Klasifikasi Kemudahan Pasar:** Pengelompokan Kabupaten/Kota ke dalam 5 kategori: *Sangat Mudah, Mudah, Moderat, Sulit,* dan *Sangat Sulit (Alot)*.
- **Segmentasi Klaster:** Pemetaan pasar ruko (*Cluster 1: Prime Commercial Urban*, *Cluster 0: Tight Market / High Landlord Power*, *Cluster 2: Fast Velocity & High Yield*).
- **Leaderboard Wilayah:** Top 10 wilayah termudah vs Top 10 wilayah paling alot dengan landlord strict.

### 4. Siklus Lead Time & SLA
- **Dekomposisi 7 Siklus:** Evaluasi lead time dari Survei, Approval Direksi, Negosiasi, Legalitas MoU, Renovasi, Rekrutmen SDM, hingga Grand Opening.
- **Identifikasi Bottleneck:** Deteksi deviasi tahap renovasi kontraktor (+4.2 hari di atas SLA).
- **Law of Diminishing Returns:** Uji statistik durasi negosiasi vs yield diskon dengan rekomendasi batas maksimal (hard stop) pada hari ke-15.

### 5. Manajemen Risiko Vendor & Renovasi
- **Kuadran Risiko:** Matriks deviasi anggaran (*cost variance*) vs keterlambatan waktu pengerjaan (*time delay*).
- **Klasifikasi Kontraktor:** Tier 1 (Mitra Unggulan), Tier 2 (Moderat), Tier 3 (Risiko Tinggi), dan Kasus Kritis (Red Flag Kontraktor Bali).

### 6. Data Explorer & Evaluasi Batch
- **Pencarian Interaktif:** Filter dinamis berdasarkan tahun, nama negosiator, wilayah, dan rentang sewa.
- **Evaluasi Massal (Batch AI):** Pemrosesan serentak file spreadsheet Excel/CSV calon ruko baru.
- **Ekspor Data:** Fitur unduh data terfilter atau hasil evaluasi langsung ke format Excel (.xlsx) dan CSV.

---

## Panduan Menjalankan Aplikasi Secara Lokal

### Prasyarat
- Python 3.10 atau versi yang lebih baru
- Git

### Langkah Instalasi
```bash
# 1. Clone repositori ini
git clone https://github.com/Mufti129/nego-intelligence.git
cd nego-intelligence

# 2. Buat dan aktifkan virtual environment (opsional)
python3 -m venv venv
source venv/bin/activate   # Di Windows: venv\Scripts\activate

# 3. Pasang dependensi
pip install -r requirements.txt

# 4. Jalankan aplikasi Streamlit
streamlit run app.py
```
Aplikasi akan aktif pada alamat: `http://localhost:8501`.

---

## Panduan Deploy ke Streamlit Community Cloud

1. Buka [share.streamlit.io](https://share.streamlit.io/) dan login menggunakan akun GitHub Anda.
2. Klik tombol **New app**.
3. Pilih repository **`Mufti129/nego-intelligence`**, branch **`main`**, dan Main file path **`app.py`**.
4. Klik **Deploy!**

---

## Struktur Berkas

```text
├── .streamlit/
│   └── config.toml             # Konfigurasi tema korporat & server
├── data_nego_baru/             # Database resmi cabang
│   └── data_nego_baru_cleaned_2024-2026.xlsx
├── app.py                      # Aplikasi utama Streamlit Dashboard
├── streamlit_app.py            # Entrypoint kompatibilitas Streamlit Cloud
├── requirements.txt            # Dependensi paket Python
├── properties_cleaned_2024-2026.csv # Database cadangan CSV
├── .gitignore                  # Pengecualian berkas sistem & binary
└── README.md                   # Dokumentasi resmi sistem
```

---

## Pengembang
- **Mukhammad Rekza Mufti** — *Data Analyst (Divisi Bisnis & Ekspansi Jaringan UPC)*
- **Pusat Gadai Indonesia (PGI)**
