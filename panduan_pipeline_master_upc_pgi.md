# PANDUAN PENGGUNAAN MASTER PIPELINE SISTEM ANALISIS UPC PGI (2024–2026)

**Skrip Master**: [`pipeline_master_analisis_upc_pgi.py`](file:///Users/macbookair/Documents/Analisa/Nego_baru/pipeline_master_analisis_upc_pgi.py)  
**Dokumen Laporan Resmi**: [`Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx`](file:///Users/macbookair/Documents/Analisa/Nego_baru/Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx)  
**Dokumentasi Metodologis**: [`DOKUMENTASI_LENGKAP_PROSES_ANALISIS_UPC_PGI.md`](file:///Users/macbookair/Documents/Analisa/Nego_baru/DOKUMENTASI_LENGKAP_PROSES_ANALISIS_UPC_PGI.md)  
**Entitas**: Pusat Gadai Indonesia (PGI) — Divisi Bisnis  

---

## 1. PENDAHULUAN & TUJUAN SISTEM

Skrip `pipeline_master_analisis_upc_pgi.py` dirancang untuk menyatukan seluruh alur kerja operasional data analisis pembukaan cabang baru (*Unit Pelayanan Cabang / UPC*) PGI ke dalam **satu sistem terpadu yang interaktif**.

Mengacu pada arsitektur di folder `Script lama`, sistem ini bekerja mulai dari data mentah (*raw data*), melakukan audit & pembersihan data (*cleansing*), hingga menyajikan **Menu Pilihan Interaktif di Terminal** di mana Anda dapat memilih analisis mana yang ingin dieksekusi sesuai bab pada laporan resmi direksi.

---

## 2. CARA MENJALANKAN DI TERMINAL VSCODE / MACOS

1. Buka folder kerja di VSCode:
   ```bash
   cd /Users/macbookair/Documents/Analisa/Nego_baru
   ```
2. Buka terminal terintegrasi (Pintasan: `` Ctrl + ` ``)
3. Jalankan perintah:
   ```bash
   python3 pipeline_master_analisis_upc_pgi.py
   ```

---

## 3. PANDUAN LENGKAP MENU & METODOLOGI ANALISIS

Setelah skrip dijalankan, akan muncul menu interaktif seperti berikut:

```text
===============================================================================================
    🚀 MASTER PIPELINE SISTEM ANALISIS DATA EKSPANSI & NEGOSIASI CABANG (UPC) PGI
===============================================================================================
  Status Database: 1028 Cabang Unik Terdaftar (2024–2026)
  Direktori Output: /Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis
-----------------------------------------------------------------------------------------------
  [PILIHAN MODUL ANALISIS & PIPELINE]
   1. Pipeline Data Cleansing & Audit Kesehatan Data (Deduplikasi Multi-Termin 1.047 -> 1.028)
   2. Dekomposisi 6 Tahapan Siklus Waktu & Harmonisasi Lead Time UPC (58 hari vs 69 hari)
   3. Evaluasi Kinerja Negosiasi & Negosiator (Bonita vs Mirza vs Tim Surveyor + Uji Mann-Whitney U)
   4. Uji Korelasi Durasi vs Diskon & Pembuktian Empiris Law of Diminishing Returns
   5. Analisis Spasial & Geospasial Kemudahan Negosiasi Wilayah (Top 10 Termudah/Tersulit + TopoJSON)
   6. Vendor Risk Quadrant Kontraktor Renovasi & Dekomposisi Kasus Khusus Bali (Trimo)
   7. Smart Auto-Routing Engine & Simulasi Interaktif Negosiator (Sistem Penugasan Ruko Baru)
   8. Kompilasi Dokumen Laporan Resmi Word (.docx) & Excel Multi-Sheet
   9. Jalankan Seluruh Pipeline Sekaligus (Batch All-in-One Execution)
   0. Keluar dari Program (Exit)
===============================================================================================
```

### Menu 1: Pipeline Data Cleansing & Audit Kesehatan Data
* **Fungsi**: Membaca data mentah `properties (1) - properties (2024-2026)-.csv`, mendeteksi 17 cabang fisik yang memiliki 19 baris pembayaran bertahap (termin renovasi), dan melakukan deduplikasi dengan aturan `keep='last'` pada tanggal serah terima final.
* **Output**: Menghasilkan file bersih terstandarisasi `properties_cleaned_2024-2026.csv` (1.028 cabang unik mandiri: 317 di 2024, 436 di 2025, 275 di 2026).

### Menu 2: Dekomposisi 6 Tahapan Siklus Waktu & Harmonisasi Lead Time UPC
* **Fungsi**: Menghitung siklus hidup 6 fase UPC:
  1. Pengajuan s/d Approval Direksi
  2. Tunggu Negosiasi
  3. Durasi Negosiasi Riil
  4. Tunggu Renovasi
  5. Durasi Renovasi Fisik Kontraktor
  6. Tunggu Grand Opening (*paid_at* s/d Buka Toko)
* **Output**: Tabel dekomposisi tahunan, penjelasan harmonisasi Lead Time Pasca-Persetujuan (58,7 hari $\approx$ 58 hari) vs Full End-to-End (69,2 hari), bukti pemotongan waktu tunggu Grand Opening sebesar 60% (14,0 hr $\rightarrow$ 9,4 hr $\rightarrow$ 5,6 hr), serta ekspor `4_dekomposisi_siklus_waktu_upc.png`.

### Menu 3: Evaluasi Kinerja Negosiasi & Negosiator (Bonita vs Mirza)
* **Fungsi**: Menilai efektivitas Tim Negosiasi Resmi (Bonita & Mirza) vs Tim Surveyor Lapangan (Dika & Salma).
* **Temuan Utama**:
  * Bonita & Mirza menangani 264 cabang (96,0%) dan menyumbang 100% saving nasional (**Rp 2,04 Miliar**).
  * 11 cabang dari tim survei ditutup di harga awal (diskon 0%, saving Rp 0).
  * Uji Statistik **Mann-Whitney U Test** ($U = 11.083, p = 0,00012$) membuktikan keunggulan Bonita pada ruko nilai tinggi (*The High-Value Negotiator*) dan Mirza pada kecepatan perputaran (*The Speed Specialist*).
* **Output**: Tabel performa bulanan, Grafik 1 (Volume & Saving), Grafik 2 (Komparasi Saving), Grafik 3 (Diskon & Durasi).

### Menu 4: Uji Korelasi Durasi vs Diskon & Law of Diminishing Returns
* **Fungsi**: Menguji apakah menunda kesepakatan lebih lama menghasilkan diskon lebih besar.
* **Temuan Utama**:
  * Korelasi Pearson bernilai negatif insignifikan ($r = -0,1040, p = 0,085$), membuktikan **TIDAK ADA hubungan positif** antara lamanya nego dengan diskon.
  * *Law of Diminishing Returns*: Transaksi di atas 20 hari mencatatkan tingkat kegagalan meraih diskon melonjak menjadi **27,5%** (11 cabang gagal dapat diskon).
  * Zona Emas adalah 1–10 hari (tingkat sukses 93%–97%).
* **Output**: Matriks bucketing durasi, Grafik 7 (`7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png`).

### Menu 5: Analisis Spasial & Geospasial Kemudahan Negosiasi Wilayah
* **Fungsi**: Mengkalkulasi Indeks Kemudahan Negosiasi Wilayah (Skor 0–100 dengan bobot seimbang 50% Kecepatan + 50% Diskon) dan memetakan 76 kabupaten/kota aktif ke TopoJSON 524 wilayah BPS.
* **Temuan Utama**:
  * Top Termudah: Cilacap (91,9), Serang (76,0), Banjarnegara (75,8), Kota Tasikmalaya (75,7), Jakarta Pusat (73,5).
  * Top Tersulit / Alot: Kab. Bandung Barat (28,6), Gianyar (30,7), Grobogan (32,9), Kota Cirebon (34,8), Cianjur (36,7).
  * Uji Empiris Kota vs Kabupaten: 8 dari 10 wilayah paling alot adalah **KABUPATEN** karena dominasi aset warisan keluarga majemuk.
* **Output**: Peta Geospasial TopoJSON Resolusi Tinggi (`5_peta_geospasial_kemudahan_nego_2026.png`) dan Analisis Klaster (`6_analisis_karakteristik_wilayah_nego_2026.png`).

### Menu 6: Vendor Risk Quadrant & Dekomposisi Kasus Khusus Bali (Trimo)
* **Fungsi**: Memetakan kontraktor rekanan renovasi ke dalam 4 Kuadran (Core Champion, Risiko Operasional, Underperformer, Potensial) serta menguji anjloknya kepatuhan SLA Trimo di tahun 2026.
* **Temuan Objektif Lapangan**:
  * Kontraktor lain (Sarjono, Edwin, Sendy, Teguh) menolak/belum bersedia mengambil proyek di Pulau Bali.
  * Trimo menyerap 92,3% proyek Bali (12 dari 13 cabang) dengan kendala penyeberangan fery Ketapang-Gilimanuk, material Jawa, dan upacara adat Banjar (durasi 43,9 hari, SLA 0%).
  * Pada proyek regular Jawa, Trimo mencatat durasi 31,8 hari (SLA 19,5%).
* **Output**: Berkas Excel multi-sheet `Laporan_Risiko_Vendor_Renovasi.xlsx`, Grafik 8 (Side-by-Side Kuadran 2025 vs 2026), Grafik 9 (Vektor Lintasan Migrasi).

### Menu 7: Smart Auto-Routing Engine & Simulasi Interaktif Negosiator
* **Fungsi**: Diadopsi langsung dari kapabilitas `Script lama`, menyediakan sistem penugasan otomatis ruko baru:
  1. *Smart Auto-Routing*: Memasukkan nama wilayah dan harga penawaran ruko, sistem secara cerdas menugaskan apakah ke Bonita (basis stronghold / ruko > Rp 50 Juta) atau Mirza (kecepatan / volume) dengan mempertimbangkan batas kapasitas pipeline aktif.
  2. *Simulasi Prediksi Kinerja*: Menghitung target diskon %, proyeksi hemat Rp, target net, dan estimasi waktu closing untuk negosiator tertentu.
  3. *Head-to-Head 37 Wilayah Overlap*: Menampilkan komparasi historis Bonita vs Mirza pada wilayah yang sama.
  4. *Batch Multi-Branch Evaluator*: Memasukkan beberapa ruko sekaligus dengan pemisah pipe (`|`) untuk memproyeksikan total penghematan batch.
  5. *Config Manager*: Mengubah SLA target days (default 17 hr), high-value limit (default Rp 50 Jt), dan kapasitas max (default 15 cabang).

### Menu 8: Kompilasi Dokumen Laporan Resmi Word (.docx)
* **Fungsi**: Mengotomatisasi penyusunan dokumen laporan Word eksekutif `Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx` (4,4 MB) lengkap dengan kop surat resmi, penomoran halaman, tabel bergaris terstandarisasi, callout catatan manajerial, dan 9 diagram 300 DPI tersemat utuh.

### Menu 9: Jalankan Seluruh Pipeline Sekaligus (Batch All-in-One)
* **Fungsi**: Melakukan eksekusi end-to-end tanpa intervensi manual (Cleansing $\rightarrow$ Modul 2 s/d 6 $\rightarrow$ Seluruh 9 Grafik $\rightarrow$ Excel $\rightarrow$ Dokumen Word).

---

## 4. DAFTAR ARTIFAK & ASET YANG DIHASILKAN

Semua berkas hasil analisis disimpan otomatis di dalam folder `hasil_analisis/`:

| Tipe Berkas | Nama Berkas | Keterangan |
| :--- | :--- | :--- |
| **Dataset Bersih** | `data_nego_baru/data_nego_baru_cleaned_2024-2026.xlsx` | 1.039 cabang unik mandiri pasca-rekonsiliasi 3 sumber |
| **Laporan Word** | `Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx` | Laporan direksi lengkap 7 Bab |
| **Laporan Excel** | `hasil_analisis/Laporan_Risiko_Vendor_Renovasi.xlsx` | Multi-sheet analisis kontraktor & kasus Bali |
| **Grafik 1** | `hasil_analisis/grafik/1_tren_volume_dan_saving_2026.png` | Tren Bulanan Volume & Total Saving 2026 |
| **Grafik 2** | `hasil_analisis/grafik/2_komparasi_saving_bonita_mirza_2026.png` | Komparasi Saving Bulanan: Bonita vs Mirza |
| **Grafik 3** | `hasil_analisis/grafik/3_komparasi_diskon_dan_durasi_2026.png` | Komparasi Diskon % & Kecepatan Durasi |
| **Grafik 4** | `hasil_analisis/grafik/4_dekomposisi_siklus_waktu_upc.png` | Bar Chart Dekomposisi Siklus Waktu Lead Time (7 Tahapan) |
| **Grafik 4a** | `hasil_analisis/grafik/4a_tabel_dekomposisi_7_siklus_waktu.png` | **BARU**: Tabeling Dekomposisi 7 Siklus Tahunan (2024–2026) |
| **Grafik 4b** | `hasil_analisis/grafik/4b_tabel_detail_bulanan_siklus_2026.png` | **BARU**: Tabeling Detail Bulanan 7 Siklus Tahun 2026 (Jan–Agu) |
| **Dashboard 4c** | `hasil_analisis/grafik/4_dekomposisi_7_siklus_tabel_komprehensif.png` | **BARU**: Dashboard Eksekutif Tabel 7 Siklus & Bulanan 2026 |
| **Grafik 5** | `hasil_analisis/grafik/5_peta_geospasial_kemudahan_nego_2026.png` | Peta Tematik TopoJSON BPS Kemudahan Nego |
| **Grafik 6** | `hasil_analisis/grafik/6_analisis_karakteristik_wilayah_nego_2026.png` | Analisis Karakteristik Top Wilayah & 11 Klaster |
| **Grafik 7** | `hasil_analisis/grafik/7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png` | Uji Korelasi & Law of Diminishing Returns |
| **Grafik 8** | `hasil_analisis/grafik/8_peta_risiko_vendor_komparasi_2025_2026.png` | Vendor Risk Quadrant Komparasi 2025 vs 2026 |
| **Grafik 8a** | `hasil_analisis/grafik/8a_tabel_matriks_komparasi_vendor_2025_2026.png` | **BARU**: Tabel Matriks Komparasi Kuantitatif Performa & Migrasi Kuadran Seluruh Kontraktor |
| **Grafik 9** | `hasil_analisis/grafik/9_vektor_migrasi_risiko_vendor_2025_2026.png` | Peta Vektor Lintasan Migrasi Risiko Kontraktor |

---

*PT Pusat Gadai Indonesia (PGI) — Business Development & Operational Data Analyst*
