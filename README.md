# 🏢 Pusat Gadai Indonesia (PGI) — Smart Negotiation & UPC Expansion Intelligence Dashboard

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

Sistem analitik terpadu korporasi berbasis kecerdasan data (Data Intelligence) untuk evaluasi kinerja negosiasi ruko cabang (UPC), dekomposisi siklus lead time, pemetaan geospasial kemudahan negosiasi wilayah, audit risiko vendor renovasi, serta **Smart Auto-Routing Engine** berbasis AI / Multi-Criteria Decision Analysis (MCDA) untuk penugasan negosiator ruko baru secara objektif, cepat, dan terukur.

---

## 🌟 Fitur Utama Aplikasi

### 1. 📊 Executive KPI & Overview Dashboard
- **Metrik Kunci:** Pelacakan total cabang (1.039 database), efisiensi diskon riil (Rp 2,13 Miliar), rata-rata yield diskon (15.60%), dan SLA durasi negosiasi (11.8 hari vs target 17 hari).
- **Komparasi Kinerja Tim:** Analisis mendalam kinerja **Bonita** (Senior Specialist: 142 deal, Rp 1,40 M saving, 18.25% diskon) vs **Mirza** (Strategic Negotiator: 132 deal, Rp 728,4 Juta saving, 14.05% diskon) vs **Tim Lapangan**.
- **Visualisasi Interaktif:** Boxplot distribusi diskon dan scatter plot harga penawaran awal vs harga deal final.

### 2. 🤖 Smart AI Auto-Routing & Live Deal Simulator
- **Multi-Criteria Decision Analysis (MCDA):** Mengintegrasikan 4 pilar utama:
  1. *Workload Balance & Fair Capacity* (Beban kerja aktif tim ≤ 15 proyek)
  2. *Regional Historical Track Record* (Stronghold historis wilayah)
  3. *High-Value Asset Affinity* (Nilai penawaran ruko ≥ Rp 50 Juta)
  4. *SLA Speed Optimization* (Target SLA ≤ 17 hari)
- **Live Output:** Rekomendasi negosiator terbaik, alasan keputusan otomatis, estimasi saving nominal (Rp), estimasi diskon (%), estimasi durasi (hari), dan Radar Chart skor 4 pilar.

### 3. 🗺️ Geospatial & Regional Intelligence
- **Klasifikasi Kemudahan Negosiasi Wilayah:** Pengelompokan seluruh Kabupaten/Kota ke dalam 5 kategori: *Sangat Mudah, Mudah, Moderat, Sulit,* dan *Sangat Sulit (Alot)*.
- **Klaster Karakteristik Pasar:** Segmentasi pasar ruko (*Cluster 1: Prime Commercial Urban*, *Cluster 0: Tight Market / High Landlord Power*, *Cluster 2: Fast Velocity & High Yield*).
- **Leaderboard Wilayah:** Top 10 wilayah termudah vs Top 10 wilayah paling alot dengan landlord strict.

### 4. ⏱️ 7-Stage Cycle Decomposition & SLA Tracking
- **Dekomposisi 7 Siklus:** Evaluasi lead time dari Survei, Approval Direksi, Negosiasi, Legalitas MoU, Renovasi, Rekrutmen SDM, hingga Grand Opening.
- **Harmonisasi Lead Time:** Identifikasi bottleneck utama pada tahap renovasi kontraktor (+4.2 hari deviasi).
- **Law of Diminishing Returns:** Pembuktian statistik bahwa negosiasi > 14 hari tidak menghasilkan kenaikan diskon yang signifikan (Rekomendasi *hard stop* hari ke-15).

### 5. 🏗️ Vendor Risk & Kontraktor Renovasi
- **Kuadran Risiko Kontraktor:** Matriks perbandingan antara deviasi biaya (*cost variance*) dan keterlambatan waktu (*time delay*).
- **Evaluasi Kontraktor:** Klasifikasi Tier 1 (Mitra Unggulan), Tier 2 (Moderat), Tier 3 (Risiko Tinggi), dan Kasus Kritis (Red Flag Kontraktor Trimo Bali).

### 6. 🔍 Batch Multi-Ruko AI Evaluator & Data Explorer
- **Eksplorasi Fleksibel:** Filter dinamis berdasarkan tahun, nama negosiator, wilayah, dan rentang nilai sewa.
- **Batch AI Evaluator:** Unggah file Excel/CSV berisi daftar calon ruko baru untuk dievaluasi oleh sistem AI secara massal dan serentak.
- **Ekspor Hasil:** Unduh data hasil filter atau hasil evaluasi AI langsung ke format Excel (.xlsx) dan CSV.

---

## 🚀 Panduan Menjalankan Aplikasi Secara Lokal

### Prasyarat
- Python 3.10 atau versi yang lebih baru
- Git

### Langkah Instalasi
```bash
# 1. Clone repositori ini
git clone <URL_REPOSITORI_ANDA>
cd <NAMA_FOLDER_REPOSITORI>

# 2. Buat dan aktifkan virtual environment (opsional namun disarankan)
python3 -m venv venv
source venv/bin/activate   # Di Windows: venv\Scripts\activate

# 3. Pasang dependensi
pip install -r requirements.txt

# 4. Jalankan aplikasi Streamlit
streamlit run app.py
```
Aplikasi akan terbuka otomatis di peramban (browser) Anda pada alamat: `http://localhost:8501`.

---

## ☁️ Panduan Deploy ke Streamlit Community Cloud (Gratis)

1. **Unggah Repositori ke GitHub:**
   - Pastikan seluruh berkas (`app.py`, `streamlit_app.py`, `requirements.txt`, `.streamlit/config.toml`, dan folder `data_nego_baru/` atau dataset `.csv`) sudah ter-push ke GitHub repository Anda.
2. **Buka Streamlit Cloud:**
   - Kunjungi [share.streamlit.io](https://share.streamlit.io/) dan masuk menggunakan akun GitHub Anda.
3. **Buat Aplikasi Baru (New App):**
   - Pilih Repositori GitHub Anda.
   - Branch: `main` (atau `master`).
   - Main file path: `app.py` atau `streamlit_app.py`.
4. **Deploy:**
   - Klik **Deploy!** Dalam 1–2 menit, dashboard Anda akan aktif secara online dengan link publik HTTPS resmi.

---

## 📂 Struktur Proyek

```text
├── .streamlit/
│   └── config.toml             # Konfigurasi tema modern & parameter server
├── data_nego_baru/             # Berkas dataset resmi
│   └── data_nego_baru_cleaned_2024-2026.xlsx
├── hasil_analisis/             # Berkas output visualisasi & grafik
├── app.py                      # Aplikasi utama Streamlit Dashboard (Full-Featured)
├── streamlit_app.py            # Entrypoint kompatibilitas Streamlit Cloud
├── requirements.txt            # Daftar paket dependensi Python
├── properties_cleaned_2024-2026.csv # Dataset cadangan CSV
├── .gitignore                  # Berkas pengecualian Git
└── README.md                   # Dokumentasi resmi proyek
```

---

## 👨‍💻 Penulis & Pengembang
- **Mukhammad Rekza Mufti** — *Data Analyst (Divisi Bisnis & Ekspansi Jaringan UPC)*
- **Pusat Gadai Indonesia (PGI)**
