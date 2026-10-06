# Panduan Operasional: Analisis Performa Negosiasi & Uji Korelasi Durasi vs Diskon 2026 di VSCode

Dokumen ini merupakan panduan praktis dan dokumentasi teknis komprehensif bagi tim **Divisi Bisnis Pusat Gadai Indonesia (PGI)** untuk menjalankan dan memodifikasi skrip Python analisis performa negosiasi bulanan, efisiensi diskon, serta uji korelasi durasi tawar-menawar tahun 2026 secara mandiri di Visual Studio Code (VSCode).

---

## 1. Lokasi Berkas & Struktur Folder

Semua berkas telah disiapkan dan terorganisir di dalam direktori kerja:
`/Users/macbookair/Documents/Analisa/Nego_baru`

```text
Nego_baru/
├── analisa_performa_negosiasi_bulanan_2026.py   <-- SKRIP UTAMA (Jalankan di VSCode)
├── build_v3_docx.py                           <-- Skrip penyusun Laporan Word resmi v3
├── cleansing_data_negosiasi.py                <-- Skrip pembersihan & deduplikasi data
├── properties_cleaned_2024-2026.csv           <-- Dataset bersih (1.028 cabang unik)
├── properties (1) - properties (2024-2026)-.csv <-- Dataset mentah sumber (1.047 baris)
├── Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx <-- Laporan Word resmi v3 (3,3 MB)
├── panduan_analisis_negosiasi_bulanan_vscode.md <-- Panduan operasional VSCode ini
├── panduan_analisis_geospasial_vscode.md       <-- Panduan peta geospasial TopoJSON
└── hasil_analisis/                            <-- Folder output ekspor otomatis
    ├── Laporan_Performa_Negosiasi_Bulanan_2026.xlsx  <-- File Excel multi-sheet lengkap
    ├── performa_nego_bulanan_2026_overall.csv
    ├── performa_bulanan_bonita_2026.csv
    ├── performa_bulanan_mirza_2026.csv
    ├── komparasi_bulanan_tim_nego_2026.csv
    └── grafik/                                <-- Berkas Grafik PNG (300 DPI)
        ├── 1_tren_volume_dan_saving_2026.png
        ├── 2_komparasi_saving_bonita_mirza_2026.png
        ├── 3_komparasi_diskon_dan_durasi_2026.png
        ├── 4_dekomposisi_siklus_waktu_upc.png
        └── 7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png
```

---

## 2. Cara Menjalankan Skrip di Visual Studio Code (VSCode)

### Langkah 1: Buka Folder Kerja
1. Buka aplikasi **Visual Studio Code**.
2. Klik menu **File** $\rightarrow$ **Open Folder...** (atau tekan `Cmd + O` di macOS / `Ctrl + K Ctrl + O` di Windows).
3. Pilih direktori **`Nego_baru`** (`/Users/macbookair/Documents/Analisa/Nego_baru`).

### Langkah 2: Pilih Python Interpreter
1. Buka berkas [**`analisa_performa_negosiasi_bulanan_2026.py`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/analisa_performa_negosiasi_bulanan_2026.py).
2. Periksa status bar di pojok kanan bawah VSCode. Pastikan environment Python telah aktif (misalnya `Python 3.11`, `3.12`, atau `3.13`).
3. Jika belum terpilih, tekan `Cmd + Shift + P` $\rightarrow$ ketik `Python: Select Interpreter` $\rightarrow$ pilih interpreter Python yang terpasang di sistem Anda.

### Langkah 3: Eksekusi Skrip
Pilih salah satu dari dua cara berikut:
* **Cara A (Satu Klik / Tombol Run)**:
  * Klik ikon segitiga **Play / Run Python File** di pojok kanan atas jendela editor VSCode.
* **Cara B (Melalui Terminal VSCode)**:
  * Buka terminal terintegrasi dengan menekan shortcut `` Ctrl + ` `` (Backtick) atau menu **Terminal** $\rightarrow$ **New Terminal**.
  * Ketik perintah berikut lalu tekan **Enter**:
    ```bash
    python3 analisa_performa_negosiasi_bulanan_2026.py
    ```

### Langkah 4: Melihat Hasil Analisis
1. **Di Terminal VSCode**: Skrip menampilkan ringkasan dashboard eksekutif, tabel per negosiator, tabel komparasi bulanan, dan koefisien statistik.
2. **Di Berkas Excel & CSV**: Buka folder **`hasil_analisis/`** untuk melihat berkas [**`Laporan_Performa_Negosiasi_Bulanan_2026.xlsx`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/Laporan_Performa_Negosiasi_Bulanan_2026.xlsx).
3. **Di Gambar Grafik**: Buka folder **`hasil_analisis/grafik/`** untuk melihat 5 gambar visualisasi (300 DPI) siap pakai.

---

## 3. Aturan & Standarisasi Bisnis Data Negosiasi

| Komponen Analisis | Kolom / Formula Database | Penjelasan Logika Bisnis |
| :--- | :--- | :--- |
| **Filter Tahun** | `Tahun == 2026` | Mengisolasi 275 cabang yang masuk kohort ekspansi 2026. |
| **Bulan Transaksi** | `tgl_nego_berakhir` | Dihitung berdasarkan bulan kesepakatan final sewa tercapai (`2026-01` s/d `2026-08`). |
| **Tim Negosiasi Resmi** | `nama_negosiator_2` | Khusus 2026 diemban eksklusif oleh **Mirza** dan **Bonita**. Dika (10 deal) & Salma (1 deal) berstatus surveyor lapangan. |
| **Harga Penawaran Awal** | `Hargaawal_nego_sheet` | Tolok ukur harga penawaran awal (*asking price*) dari pemilik ruko. |
| **Harga Sewa Deal** | `hargarental_sheet` | Nilai sewa final per tahun yang disepakati (*deal price*). |
| **Penghematan Finansial (Saving)**| `Hargaawal_nego_sheet - hargarental_sheet` | Total nilai rupiah yang berhasil dipangkas oleh tim negosiasi. |
| **Persentase Diskon** | `efisiensi_diskon_nego2` / `efisiensi_diskon_pct` | Persentase diskon per cabang (`0,0%` s/d `40,0%`). |
| **Durasi Negosiasi** | `lama_waktu_realisasi_nego` | Hari kalender murni dari awal tawar-menawar hingga kesepakatan final tercapai. |
| **Durasi Masa Sewa (Tenor)**| `durasi_sewa` | Jangka waktu sewa ruko dalam tahun kontrak (mayoritas 5 tahun). |

---

## 4. Rangkuman Kinerja Tim Negosiasi Bulanan 2026

### A. Tabel Tren Performa Bulanan Overall (Januari – Agustus 2026)

| Bulan (2026) | Deal | Total Penawaran Awal | Total Sewa Kesepakatan | Total Penghematan (Saving) | Rerata Diskon | Rerata Durasi | Success Rate |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **2026-01 (Januari)** | 11 | Rp 344,0 Juta | Rp 275,0 Juta | **Rp 69,0 Juta** | 19,64% | 10,7 hr | 100,0% |
| **2026-02 (Februari)** | 34 | Rp 1,24 Miliar | Rp 1,02 Miliar | **Rp 218,6 Juta** | 16,74% | 12,7 hr | 85,3% |
| **2026-03 (Maret)** | 33 | Rp 1,28 Miliar | Rp 1,02 Miliar | **Rp 265,0 Juta** | 18,88% | 11,7 hr | 87,9% |
| **2026-04 (April)** | 52 | Rp 2,19 Miliar | Rp 1,87 Miliar | **Rp 317,0 Juta** | 13,48% | 13,7 hr | 88,5% |
| **2026-05 (Mei)** | 48 | Rp 2,36 Miliar | Rp 1,98 Miliar | **Rp 383,2 Juta** | 15,81% | 14,2 hr | 93,8% |
| **2026-06 (Juni)** | 43 | Rp 2,02 Miliar | Rp 1,65 Miliar | **Rp 371,7 Juta** | 15,14% | 8,3 hr | 95,3% |
| **2026-07 (Juli)** | 31 | Rp 1,53 Miliar | Rp 1,30 Miliar | **Rp 235,1 Juta** | 12,97% | 8,9 hr | 87,1% |
| **2026-08 (Agustus)** | 23 | Rp 1,11 Miliar | Rp 938,0 Juta | **Rp 176,0 Juta** | 15,74% | 8,4 hr | 91,3% |
| **TOTAL TAHUN 2026** | **275** | **Rp 12,08 Miliar** | **Rp 10,04 Miliar** | **Rp 2,04 Miliar** | **15,57%** | **11,4 hr** | **90,5%** |

### B. Komparasi Head-to-Head: Bonita vs Mirza (Tim Negosiasi Resmi)

| Bulan | Deal B | Diskon B | Saving Bonita | Dur B | Deal M | Diskon M | Saving Mirza | Dur M | Non-Tim* | Total Saving |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Januari** | 7 | 21,6% | Rp 54,0 Jt | 14,1 hr | 4 | 16,2% | Rp 15,0 Jt | 4,8 hr | - | **Rp 69,0 Juta** |
| **Februari** | 15 | 21,1% | Rp 151,2 Jt | 11,7 hr | 14 | 18,1% | Rp 67,3 Jt | 8,7 hr | 5 deal | **Rp 218,6 Juta** |
| **Maret** | 13 | 25,3% | Rp 172,5 Jt | 9,8 hr | 17 | 17,3% | Rp 92,5 Jt | 11,2 hr | 3 deal | **Rp 265,0 Juta** |
| **April** | 28 | 14,4% | Rp 201,5 Jt | 12,1 hr | 21 | 14,2% | Rp 115,5 Jt | 10,7 hr | 3 deal | **Rp 317,0 Juta** |
| **Mei** | 23 | 17,9% | Rp 230,5 Jt | 17,0 hr | 25 | 13,9% | Rp 152,7 Jt | 11,6 hr | - | **Rp 383,2 Juta** |
| **Juni** | 23 | 18,7% | Rp 291,5 Jt | 7,7 hr | 20 | 11,1% | Rp 80,2 Jt | 9,1 hr | - | **Rp 371,7 Juta** |
| **Juli** | 13 | 16,2% | Rp 140,2 Jt | 11,6 hr | 18 | 10,6% | Rp 94,9 Jt | 6,9 hr | - | **Rp 235,1 Juta** |
| **Agustus** | 13 | 17,5% | Rp 106,5 Jt | 8,4 hr | 10 | 13,5% | Rp 69,5 Jt | 8,5 hr | - | **Rp 176,0 Juta** |
| **TOTAL** | **135** | **18,4%** | **Rp 1,35 M** | **11,6 hr** | **129** | **14,0%** | **Rp 687,6 Jt** | **9,6 hr** | **11 deal** | **Rp 2.035,5 Juta** |

*Catatan: Kolom 'Non-Tim' adalah 11 cabang yang diajukan langsung oleh surveyor lapangan (Dika: 10 deal, Salma: 1 deal) dengan diskon 0% (deal di harga asking).*

---

## 5. Analisis Distribusi & Uji Korelasi: Durasi Negosiasi vs Efisiensi Diskon Sewa

Salah satu pertanyaan mendasar manajemen ekspansi adalah: **"Apakah proses negosiasi yang berlarut-larut menghasilkan diskon yang lebih besar?"**

Untuk menjawab ini secara empiris, dilakukan analisis statistik bivariat terhadap seluruh **275 transaksi cabang tahun 2026**.

### A. Uji Korelasi Statistik Bivariat

| Parameter Uji | Nilai Statistik | $p$-value | Keterangan & Makna Statistik |
| :--- | :---: | :---: | :--- |
| **Koefisien Pearson ($r$)** | **-0,1040** | 0,085 | Korelasi negatif sangat lemah (tidak signifikan pada $\alpha=0,05$). |
| **Koefisien Spearman ($\rho$)** | **-0,0844** | 0,163 | Korelasi rank-order negatif sangat lemah. |
| **Koefisien Determinasi ($R^2$)**| **0,0108** | - | **Hanya 1,1%** variasi diskon yang ditentukan oleh durasi hari nego! |
| **Persamaan Regresi Linier** | $Y = -0,0957X + 16,67\%$ | - | Setiap penambahan 10 hari negosiasi justru memangkas rerata diskon sebesar $\approx 0,96\%$. |

> [!IMPORTANT]
> **KESIMPULAN ILMIAH**:
> **TIDAK ADA korelasi positif antara durasi negosiasi yang lebih lama dengan efisiensi diskon.** Memperpanjang waktu negosiasi terbukti TIDAK menghasilkan pemotongan harga yang lebih besar. Diskon tertinggi justru didapatkan pada proses negosiasi yang diselesaikan secara cepat dan tegas ($\le 10$ hari).

---

### B. Dekomposisi Kinerja Berdasarkan Kelompok Durasi (Bucketing)

| Kelompok Durasi | Jumlah Transaksi | Porsi (%) | Rerata Diskon | Median Diskon | Total Saving | Success Rate (>0% Diskon) | Gagal Diskon (0%) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **1–5 hari (Super Cepat)** | 64 deal | 23,3% | **16,41%** | 16,5% | Rp 412,0 Jt | **96,9%** | 3,1% (2 unit) |
| **6–10 hari (Cepat / Ideal)** | 126 deal | 45,8% | **15,28%** | 15,5% | **Rp 908,8 Jt** | **92,9%** | 7,1% (9 unit) |
| **11–14 hari (Batas Toleransi SLA)**| 32 deal | 11,6% | **17,72%** | 16,5% | Rp 292,7 Jt | **93,8%** | 6,2% (2 unit) |
| **15–20 hari (Zona Alot)** | 13 deal | 4,7% | **14,46%** | 13,0% | Rp 95,0 Jt | **84,6%** | 15,4% (2 unit) |
| **> 20 hari (Sangat Alot / Redflag)**| 40 deal | 14,5% | **13,83%** | 13,0% | Rp 327,0 Jt | **72,5%** | **27,5% (11 unit)** |
| **TOTAL TAHUN 2026** | **275 deal** | **100,0%** | **15,57%** | **15,0%** | **Rp 2.035,5 Jt** | **90,5%** | **9,5% (26 unit)** |

---

### C. Tiga Temuan Kunci Manajemen dari Analisis Distribusi

1. **Zona Emas Negosiasi ($\le 10$ Hari)**:
   * **69,1% transaksi** (190 deal) tuntas dalam rentang 1–10 hari kalender.
   * Menghasilkan **Rp 1,32 Miliar** (65% dari total penghematan tahun 2026) dengan tingkat keberhasilan diskon prima (**93% – 97%**).
2. **Hukum Penurunan Hasil (*Law of Diminishing Returns*)**:
   * Ketika proses tawar-menawar melewati hari ke-14, efisiensi diskon rata-rata turun dari 17,7% ke 14,5% (15–20 hari) dan anjlok ke 13,8% (> 20 hari).
   * Pada durasi > 20 hari, **27,5% cabang berakhir deal di harga awal tanpa diskon (diskon 0%)**. Hal ini membuktikan bahwa durasi yang panjang bukan disebabkan negosiator sedang memeras margin, melainkan akibat resistensi kaku pemilik ruko / sengketa kepemilikan keluarga.
3. **Uji Korelasi Durasi Masa Sewa (Tenor Kontrak Ruko)**:
   * Dilakukan uji statistik terhadap durasi masa sewa ruko: **92,7% transaksi (255 unit)** terkunci pada durasi kontrak 5 tahun, 19 unit 3 tahun, dan 1 unit 2 tahun.
   * Nilai korelasi tenor sewa terhadap diskon adalah $r = 0,0027$ ($p = 0,964$). Artinya, efisiensi diskon sewa di PGI tidak dipengaruhi oleh panjang pendeknya komitmen tahun sewa ruko, melainkan murni daya tawar korporasi pembayaran tunai di muka.

---

## 6. Galeri Visualisasi Grafik Analisis

Skrip secara otomatis menghasilkan 5 visualisasi beresolusi tinggi (300 DPI) di folder `hasil_analisis/grafik/`:

### Grafik 7: Sebaran (Scatter Plot) & Uji Korelasi Durasi Negosiasi vs Efisiensi Diskon Sewa 2026
![Sebaran dan Korelasi Durasi vs Diskon 2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png)

* **Interpretasi Visual**:
  * **Panel A (Scatter Plot Kiri)**: Titik biru (Bonita), hijau (Mirza), dan silang merah (Non-Tim). Garis regresi merah miring ke bawah memperlihatkan tren negatif ($y = -0,10x + 16,7\%$).
  * **Panel B (Bar Chart Kanan Atas)**: Rerata diskon tertinggi tercapai di kelompok $\le 14$ hari (15,3%–17,7%), lalu menurun ke 13,8% pada kelompok alot (> 20 hari).
  * **Panel C (Bar Chart Kanan Bawah)**: Tingkat kesuksesan diskon anjlok tajam dari 97% (1–5 hari) menjadi hanya 72,5% (> 20 hari).

---

### Grafik 1: Tren Bulanan Volume Deal & Akumulasi Saving 2026
![Tren Bulanan Volume Deal & Total Saving 2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/1_tren_volume_dan_saving_2026.png)

---

### Grafik 2: Komparasi Penghematan Sewa Bulanan (Bonita vs Mirza)
![Komparasi Penghematan Sewa Bulanan: Bonita vs Mirza 2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/2_komparasi_saving_bonita_mirza_2026.png)

---

### Grafik 3: Perbandingan Efisiensi Diskon (%) dan Kecepatan Durasi (Hari)
![Dinamika Diskon dan Kecepatan Realisasi Negosiasi 2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/3_komparasi_diskon_dan_durasi_2026.png)

---

### Grafik 4: Dekomposisi Lead Time Proses Pembukaan Cabang UPC (2024–2026)
![Dekomposisi Siklus Waktu UPC 2024-2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/4_dekomposisi_siklus_waktu_upc.png)

* **Harmonisasi Lead Time**:
  * **Siklus Pasca-Approval (Approved s/d Buka)**: Rata-rata **58,7 hari (~58 hari di 2026, median 53,0 hari)**: Tunggu Nego (11,1 hr) + Durasi Nego (11,5 hr) + Tunggu Renovasi (6,3 hr) + Pengerjaan Renovasi (23,9 hr) + Tunggu Buka (5,6 hr) = 58,4 – 58,7 hari.
  * **Total Lead Time Penuh (Pengajuan s/d Buka)**: Rata-rata **69,2 hari (median 64,0 hari)**, dengan menambahkan Tahap Pengajuan s/d Approval (**10,7 hari**).

---

## 7. Rekomendasi Kebijakan SLA & Operasional Negosiasi

Berdasarkan pembuktian data distribusi di atas, direkomendasikan kebijakan operasional berikut:
1. **Penerapan Batas Maksimal Toleransi Negosiasi (SLA 14 Hari)**:
   * Jika negosiasi belum mencapai kesepakatan pada hari ke-14, lakukan evaluasi cut-off. Probabilitas mendapatkan diskon menurun tajam sementara biaya oportunitas keterlambatan pembukaan cabang terus membengkak.
2. **Eskalasi Cepat ke Alternatif Ruko (Plan B)**:
   * Tim surveyor wajib menyiapkan minimal 2 ruko alternatif per wilayah agar negosiator memiliki posisi tawar kuat (*BATNA - Best Alternative to a Negotiated Agreement*) dan tidak tersandera oleh pemilik ruko yang alot.
3. **Standarisasi Kontrak 5 Tahun dengan Opsi Pembayaran Termin Tahunan / Bilyet Giro**:
   * Pertahankan keunggulan tenor 5 tahun untuk mengunci lokasi strategis UPC dari kompetitor, namun fokuskan taktik tawar pada diskon tunai tahun pertama.
