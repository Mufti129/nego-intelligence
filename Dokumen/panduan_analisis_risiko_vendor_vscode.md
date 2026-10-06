# Panduan Operasional: Analisis Risiko Vendor & Dinamika Kuadran Kontraktor (2025 vs 2026) di VSCode

Dokumen ini merupakan panduan praktis dan dokumentasi teknis bagi tim **Divisi Bisnis Pusat Gadai Indonesia (PGI)** untuk menjalankan skrip Python analisis performa kontraktor renovasi ruko, pemetaan kuadran risiko vendor (*Vendor Risk Quadrant* Poin A), pelacakan vektor migrasi 2025 $\rightarrow$ 2026, serta dekomposisi dampak penugasan wilayah khusus (Provinsi Bali) secara mandiri di Visual Studio Code (VSCode).

---

## 1. Lokasi Berkas & Struktur Folder Kerja

Semua berkas terkait analisis risiko vendor berada di dalam folder kerja:
`/Users/macbookair/Documents/Analisa/Nego_baru`

```text
Nego_baru/
├── analisa_risiko_vendor.py                    <-- SKRIP UTAMA VENDOR (Jalankan di VSCode)
├── build_v3_docx.py                           <-- Skrip penyusun Laporan Word resmi v3
├── properties_cleaned_2024-2026.csv           <-- Dataset bersih (1.028 cabang unik)
├── DOOM_UPC_Operational_Data_Analyst.pdf      <-- Acuan audit operasional (Poin A: Peta Risiko)
├── Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx <-- Laporan Word resmi v3 (4,4 MB)
├── panduan_analisis_risiko_vendor_vscode.md   <-- Panduan operasional vendor ini
└── hasil_analisis/                            <-- Folder output ekspor otomatis
    ├── Laporan_Risiko_Vendor_Renovasi.xlsx    <-- Berkas Excel multi-sheet analisis vendor
    └── grafik/                                <-- Berkas Grafik PNG (300 DPI)
        ├── 8_peta_risiko_vendor_komparasi_2025_2026.png <-- Peta Kuadran Berdampingan (Side-by-Side)
        └── 9_vektor_migrasi_risiko_vendor_2025_2026.png  <-- Peta Lintasan & Vektor Migrasi 2025->2026
```

---

## 2. Cara Menjalankan Skrip di Visual Studio Code (VSCode)

### Langkah 1: Buka Folder Kerja
1. Buka aplikasi **Visual Studio Code**.
2. Klik menu **File** $\rightarrow$ **Open Folder...** (atau tekan `Cmd + O` di macOS / `Ctrl + K Ctrl + O` di Windows).
3. Pilih folder **`Nego_baru`** (`/Users/macbookair/Documents/Analisa/Nego_baru`).

### Langkah 2: Buka dan Jalankan Skrip
Pilih salah satu dari 2 cara praktis:
* **Cara A (Satu Klik / Tombol Run)**:
  * Buka berkas [**`analisa_risiko_vendor.py`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/analisa_risiko_vendor.py).
  * Klik ikon segitiga **Play / Run Python File** di pojok kanan atas jendela editor VSCode.
* **Cara B (Melalui Terminal Terintegrasi VSCode)**:
  * Buka terminal dengan shortcut `` Ctrl + ` `` (Backtick).
  * Ketik perintah berikut lalu tekan **Enter**:
    ```bash
    python3 analisa_risiko_vendor.py
    ```

### Langkah 3: Memeriksa Hasil Output
1. **Di Terminal VSCode**: Skrip menampilkan dashboard matriks komparasi 2025 vs 2026, dekomposisi Trimo Bali vs Jawa, dan rangkuman rekomendasi manajerial.
2. **Di Berkas Excel Multi-Sheet**: Buka [**`hasil_analisis/Laporan_Risiko_Vendor_Renovasi.xlsx`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/Laporan_Risiko_Vendor_Renovasi.xlsx) yang terbagi menjadi:
   * *Sheet 1: Komparasi_2025_2026* (Daftar seluruh vendor, perubahan volume, SLA, dan pergeseran kuadran).
   * *Sheet 2: Vendor_2026_Detail* (Data rincian metrik kontraktor kohort 2026).
   * *Sheet 3: Vendor_2025_Detail* (Data rincian metrik kontraktor kohort 2025).
   * *Sheet 4: Analisis_Wilayah_Bali* (Komparasi performa regional di Provinsi Bali vs Non-Bali).
3. **Di Berkas Visualisasi (300 DPI)**:
   * [**`hasil_analisis/grafik/8_peta_risiko_vendor_komparasi_2025_2026.png`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/grafik/8_peta_risiko_vendor_komparasi_2025_2026.png)
   * [**`hasil_analisis/grafik/9_vektor_migrasi_risiko_vendor_2025_2026.png`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/grafik/9_vektor_migrasi_risiko_vendor_2025_2026.png)

---

## 3. Matriks Performa & Migrasi Vendor (Tahun 2025 vs 2026)

Berdasarkan dataset bersih 1.028 cabang, berikut adalah perbandingan performa kontraktor antara Tahun 2025 (436 cabang) dan Tahun 2026 (275 cabang):

| Nama Kontraktor | Volume '25 | SLA '25 | Volume '26 | SLA '26 | Delta SLA | Durasi '26 | Beban Bali '26 | Status Migrasi Kuadran |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **Sarjono** | 109 unit | 88,1% | 73 unit | **89,0%** | **+1,0 pt** | **14,9 hari** | 0 unit | **Kuadran I $\rightarrow$ Kuadran I (Stable Champion)** |
| **CV. Rizki Mitra Abadi** | 74 unit | 44,6% | 54 unit | **66,7%** | **+22,1 pt** | **21,7 hari** | 0 unit | **Kuadran II $\rightarrow$ Mendekati Kuadran I (Turnaround)** |
| **Edwin** | 66 unit | 69,7% | 52 unit | **61,5%** | -8,2 pt | 21,4 hari | 0 unit | Kuadran II $\rightarrow$ Kuadran II (Moderat) |
| **Trimo\*** | 74 unit | 41,9% | 53 unit | **15,1%** | **-26,8 pt** | **34,5 hari** | **12 unit (92,3%)** | Kuadran II $\rightarrow$ Kuadran II (Beban Bali) |
| **Teguh Karyanto** | 2 unit | 0,0% | 30 unit | 20,0% | +20,0 pt | 31,3 hari | 0 unit | Kuadran III $\rightarrow$ Kuadran III (Alot/Overload) |
| **Sendy** | 70 unit | 98,6% | 3 unit | 66,7% | -31,9 pt | 18,3 hari | 0 unit | Kuadran I $\rightarrow$ Kuadran III (Volume Turun Tajam) |
| **Dimas Andri Sulistyo** | - | - | 9 unit | 11,1% | Baru | 30,4 hari | 0 unit | Baru $\rightarrow$ Kuadran III (Underperformer) |
| **CV Cahaya Kemakmuran** | 39 unit | 15,4% | 0 unit | - | Off | - | 0 unit | Kuadran III $\rightarrow$ Dinonaktifkan |

---

## 4. Visualisasi Kuadran Risiko Vendor (300 DPI)

### Visualisasi 1: Peta Risiko Vendor Berdampingan (Side-by-Side 2025 vs 2026)
Menampilkan komparasi murni Poin A tanpa distraksi kartu pengujian di luar kuadran.

![Peta Risiko Vendor Komparatif](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/grafik/8_peta_risiko_vendor_komparasi_2025_2026.png)

### Visualisasi 2: Peta Vektor Lintasan Migrasi Risiko Vendor (2025 $\rightarrow$ 2026)
Melacak arah pergerakan setiap kontraktor dari posisi awal tahun 2025 (lingkaran putih putus-putus) menuju realisasi tahun 2026 (titik solid berpanah).

![Vektor Migrasi Risiko Vendor](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/grafik/9_vektor_migrasi_risiko_vendor_2025_2026.png)

---

## 5. Realitas Lapangan: Beban Penugasan Khusus Wilayah Bali pada Trimo

Salah satu temuan operasional paling penting dari audit ini adalah mengungkap penyebab utama di balik anjloknya kepatuhan SLA kontraktor **Trimo** menjadi 15,1% dan durasi pengerjaan yang membengkak ke 34,5 hari pada tahun 2026:

### Tabel Dekomposisi Performa Trimo: Bali vs Luar Bali (Tahun 2026)

| Segmentasi Wilayah Trimo 2026 | Jumlah Cabang | Rerata Durasi | Median Durasi | Kepatuhan SLA | Keterangan & Faktor Lapangan |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **Proyek Khusus Wilayah Bali** | **12 cabang (22,6%)** | **43,9 hari** | **41,5 hari** | **0,0% (0 tercapai)** | Antrean fery Ketapang-Gilimanuk, pasokan antarpulau, regulasi Banjar adat |
| **Proyek Regular Luar Bali (Jawa)** | **41 cabang (77,4%)** | **31,8 hari** | **32,0 hari** | **19,5% (8 tercapai)** | Kapasitas mandor dan perputaran modal kerja internal |
| **TOTAL PORTOFOLIO TRIMO 2026** | **53 cabang (100%)** | **34,5 hari** | **35,0 hari** | **15,1% (8 tercapai)** | **SLA agregat terdistorsi tajam oleh 12 proyek Bali** |

> [!IMPORTANT]
> **TEMUAN OPERASIONAL KRUSIAL**:
> 1. **Kontraktor Lain Menolak Wilayah Bali**: Kontraktor andalan seperti Sarjono, Sendy, Edwin, dan Teguh Karyanto **belum bersedia mengambil proyek di Provinsi Bali** karena pertimbangan jarak, kesulitan supervisi lintas selat, dan ketiadaan jaringan tukang lokal.
> 2. **Trimo Menjadi 'Bumper' Ekspansi**: Trimo menanggung **92,3% dari seluruh pembukaan cabang PGI di Bali tahun 2026** (12 dari 13 cabang Bali).
> 3. **Hambatan Lintas Selat & Adat**: Logistik material branding khusus dan brankas gadai wajib menyeberang via fery Ketapang–Gilimanuk (tambah waktu 7–14 hari). Di sisi lain, upacara keagamaan adat Banjar (Rahinan, Piodalan) mewajibkan penghentian renovasi fisik sementara.
> 4. **Evaluasi Berimbang**: Penurunan performa Trimo tidak boleh disimpulkan semata-mata sebagai ketidakmampuan manajemen, melainkan karena Trimo berani mengambil risiko ekspansi luar pulau yang dihindari oleh kontraktor lainnya.

---

## 6. Rekomendasi Strategis Manajerial

1. **Pemberlakuan SLA Regional Khusus Bali (35–40 Hari Kalender)**:
   * Target SLA renovasi nasional (21 hari kerja) tidak adil diterapkan di wilayah kepulauan seperti Bali. Terbitkan SK target SLA khusus Bali sebesar 35–40 hari kalender.
2. **Skema Insentif Mobilisasi & Berbagi Beban (*Burden Sharing*)**:
   * Jangan biarkan Trimo menanggung proyek Bali sendirian. Berikan insentif biaya mobilisasi lintas pulau bagi **Sarjono** dan **CV Rizki Mitra Abadi** agar bersedia membuka proyek di Bali.
3. **Penyediaan Hub Material Lokal di Denpasar**:
   * Jalin kemitraan langsung dengan distributor material bangunan lokal di Denpasar untuk memangkas ketergantungan pasokan fery dari Jawa Timur.
4. **Redistribusi Proyek di Pulau Jawa**:
   * Untuk proyek baru di Pulau Jawa, prioritaskan kuota utama kepada **Sarjono** (Core Champion) dan **CV Rizki Mitra Abadi** (Turnaround Improver) yang terbukti memiliki kapasitas tepat waktu.
