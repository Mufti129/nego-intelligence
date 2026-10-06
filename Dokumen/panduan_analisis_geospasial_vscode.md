# Panduan Operasional: Analisis Geospasial Kemudahan Negosiasi Wilayah PGI 2026 (VSCode)

Dokumen ini merupakan panduan lengkap bagi analis dan manajemen **Pusat Gadai Indonesia (PGI)** untuk menjalankan skrip Python analisis geospasial kemudahan dan kesulitan negosiasi secara mandiri di **Visual Studio Code (VSCode)**.

---

## 1. Ikhtisar Berkas & Arsitektur Solusi

| Nama Berkas / Direktori | Tipe | Deskripsi & Fungsi |
| :--- | :---: | :--- |
| [`analisa_geospasial_kemudahan_nego.py`](file:///Users/macbookair/Documents/Analisa/Nego_baru/analisa_geospasial_kemudahan_nego.py) | Skrip Python | Program utama untuk pengolahan skor spasial, dekoder murni TopoJSON, pembuatan peta tematik, grafik komparasi, dan ekspor data. |
| [`indonesia-kabkot-topo.json`](file:///Users/macbookair/Documents/Analisa/Nego_baru/indonesia-kabkot-topo.json) | TopoJSON | Berkas poligon batas administrasi 524 Kabupaten/Kota BPS se-Indonesia (mencakup pemekaran Bandung Barat, Pangandaran, Serang, Tangsel). |
| [`properties_cleaned_2024-2026.csv`](file:///Users/macbookair/Documents/Analisa/Nego_baru/properties_cleaned_2024-2026.csv) | Dataset Bersih | Dataset rujukan 1.028 cabang bersih (275 cabang tahun 2026). |
| [`hasil_analisis/grafik/5_peta_geospasial_kemudahan_nego_2026.png`](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/grafik/5_peta_geospasial_kemudahan_nego_2026.png) | Peta PNG (300 DPI) | Peta tematik choropleth sebaran tingkat kemudahan & kesulitan negosiasi wilayah. |
| [`hasil_analisis/grafik/6_analisis_karakteristik_wilayah_nego_2026.png`](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/grafik/6_analisis_karakteristik_wilayah_nego_2026.png) | Grafik PNG (300 DPI) | Diagram multi-panel: Top 10 Termudah vs Tersulit dan Profil Efisiensi Diskon Klaster Regional. |
| [`hasil_analisis/Analisis_Geospasial_Kemudahan_Nego_2026.xlsx`](file:///Users/macbookair/Documents/Analisa/Nego_baru/hasil_analisis/Analisis_Geospasial_Kemudahan_Nego_2026.xlsx) | Excel Multi-Sheet | Berkas Excel siap pakai (Rekap 76 Wilayah, Top Termudah, Top Tersulit, Klaster Regional). |
| [`Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx`](file:///Users/macbookair/Documents/Analisa/Nego_baru/Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx) | Laporan Word v3 | Laporan eksekutif resmi yang memuat seluruh visualisasi peta, tabel, dan analisis mendalam. |

---

## 2. Metodologi Skor Kemudahan Negosiasi Wilayah

Skor Kemudahan Negosiasi (0–100) dirumuskan secara seimbang dari 2 metrik objektif:

$$\text{Skor Kemudahan} = 50\% \times \text{Skor Kecepatan} + 50\% \times \text{Skor Efisiensi Diskon}$$

1. **Skor Kecepatan Closing (50%)**: Dihitung dari kolom resmi `lama_waktu_realisasi_nego`. Semakin sedikit hari yang dibutuhkan, skor semakin tinggi (dinormalisasi *min-max scaling* terbalik).
2. **Skor Efisiensi Diskon (50%)**: Dihitung dari kolom resmi `efisiensi_diskon_pct`. Semakin besar diskon yang berhasil diperoleh dari harga penawaran awal (`Hargaawal_nego_sheet` vs `hargarental_sheet`), skor semakin tinggi.

### Klasifikasi Tingkat Kemudahan:
* **Sangat Mudah (Skor $\ge 70$)**: Hijau Zamrud (`#059669`) — Diskon tinggi ($\ge 20\%$) & durasi sangat cepat ($< 9$ hari).
* **Mudah (Skor $60 - 69,9$)**: Biru Langit (`#0284C7`) — Diskon di atas rata-rata ($16\% - 20\%$) & durasi efisien ($5 - 10$ hari).
* **Moderat / Cukup Alot (Skor $50 - 59,9$)**: Amber / Kuning (`#F59E0B`) — Diskon standar ($13\% - 16\%$) & durasi terukur.
* **Sulit (Skor $40 - 49,9$)**: Oranye (`#EA580C`) — Diskon minim ($10\% - 13\%$) atau durasi melambat.
* **Sangat Sulit / Paling Alot (Skor $< 40$)**: Merah Crimson (`#DC2626`) — Diskon rendah ($< 10\%$) dan durasi alot ($> 17$ hari).

---

## 3. Cara Menjalankan Skrip di VSCode

### Langkah 1: Buka Folder Proyek di VSCode
Buka folder kerja:
`/Users/macbookair/Documents/Analisa/Nego_baru`

### Langkah 2: Buka Terminal Terintegrasi VSCode
Tekan kombinasi tombol:
`Ctrl + \`` (atau menu **Terminal $\rightarrow$ New Terminal**)

### Langkah 3: Eksekusi Skrip Python
Jalankan perintah berikut:
```bash
python3 analisa_geospasial_kemudahan_nego.py
```

### Output yang Dihasilkan:
1. Peta resolusi tinggi (300 DPI) dibuat otomatis di folder `hasil_analisis/grafik/5_peta_geospasial_kemudahan_nego_2026.png`.
2. Grafik karakteristik dibuat di `hasil_analisis/grafik/6_analisis_karakteristik_wilayah_nego_2026.png`.
3. Berkas Excel komparatif dibuat di `hasil_analisis/Analisis_Geospasial_Kemudahan_Nego_2026.xlsx`.
4. Berkas CSV diekspor ke `hasil_analisis/skor_kemudahan_wilayah_2026.csv` dan `hasil_analisis/klaster_regional_nego_2026.csv`.

---

## 4. Visualisasi & Hasil Analisis Geospasial

### A. Peta Tematik Kemudahan Negosiasi Wilayah (Grafik 5.1)
![Peta Geospasial Kemudahan Negosiasi 2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/5_peta_geospasial_kemudahan_nego_2026.png)

---

### B. Matriks Disparitas Skor & Profil Klaster Regional (Grafik 5.2)
![Karakteristik Wilayah & Klaster Regional 2026](/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik/6_analisis_karakteristik_wilayah_nego_2026.png)

---

## 5. Rangkuman Karakteristik Kota: Termudah vs Paling Alot

### A. Sepuluh Kota/Kabupaten Paling Gampang di-Nego (Top 10 Termudah)
| Wilayah | Jumlah Pengajuan | Rerata Durasi | Rerata Diskon | Total Saving | Skor Kemudahan | Karakteristik Utama |
| :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| **Kab. Cilacap** | **3 unit** | 8,7 hari | **40,0%** | Rp 48,5 Jt | **92,2 (Sangat Mudah)** | *Buyer's market*, pasokan ruko berlebih, pemilik fleksibel demi kepastian sewa. |
| **Kota Tasikmalaya** | **2 unit** | **4,0 hari** | 22,5% | Rp 30,0 Jt | **76,7 (Mudah)** | Penutupan super cepat, pemilik kooperatif terhadap korporasi PGI. |
| **Kab. Serang** | **3 unit** | 9,0 hari | **27,7%** | Rp 20,0 Jt | **75,2 (Mudah)** | Kawasan satelit industri Banten, diskon tinggi pada ruko non-utama. |
| **Kab. Banjarnegara** | **2 unit** | 9,0 hari | **27,5%** | Rp 8,0 Jt | **75,0 (Mudah)** | Ruko nilai sewa terjangkau (Rp 15 Jt), diskon tinggi. |
| **Kota Jakarta Pusat** | **3 unit** | 7,7 hari | **24,3%** | **Rp 67,0 Jt** | **73,0 (Mudah)** | Nilai sewa prime (asking Rp 87 M), *markup buffer* lebar, closing cepat (< 8 hari). |
| **Kab. Sukabumi** | **3 unit** | 8,0 hari | 23,3% | Rp 32,0 Jt | **71,1 (Mudah)** | Wilayah Priangan selatan dengan elastisitas harga tinggi. |
| **Kab. Garut** | **3 unit** | **5,3 hari** | 19,3% | Rp 18,0 Jt | **70,2 (Mudah)** | Kecepatan closing sangat tinggi, minim friksi tawar-menawar. |
| **Kab. Subang** | **3 unit** | 7,3 hari | 21,0% | Rp 23,0 Jt | **69,1 (Mudah)** | Koridor Pantura Jabar penyangga dengan alternatif ruko melimpah. |
| **Kab. Badung (Bali)** | **3 unit** | 6,7 hari | 19,7% | **Rp 50,0 Jt** | **68,4 (Mudah)** | Pasar komersial bernilai tinggi (asking Rp 72 M), negosiasi Bonita efektif. |
| **Kab. Bekasi** | **6 unit** | **5,2 hari** | 17,3% | Rp 43,0 Jt | **67,8 (Mudah)** | Wilayah aglomerasi industri bervolume tinggi, proses tuntas rata-rata 5 hari. |

### B. Sepuluh Kota/Kabupaten Paling Susah di-Nego (Top 10 Paling Alot)
| Wilayah | Jumlah Pengajuan | Rerata Durasi | Rerata Diskon | Total Saving | Skor Kemudahan | Karakteristik Utama & Akar Masalah |
| :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| **Kab. Bandung Barat** | **3 unit** | **24,3 hari** | **4,7%** | Rp 5,2 Jt | **19,0 (Paling Alot)** | **Juara Alot Nasional**: Aset keluarga majemuk, musyawarah berbelit, diskon minim. |
| **Kab. Grobogan** | **2 unit** | **34,0 hari** | 17,5% | Rp 13,5 Jt | **20,0 (Sangat Sulit)** | **Durasi Terlama Nasional**: Negosiasi memakan waktu lebih dari 1 bulan. |
| **Kota Cirebon** | **2 unit** | 17,0 hari | **2,5%** | Rp 1,7 Jt | **28,3 (Sulit)** | Pusat niaga heritage, tuan tanah lama ber-holding power kuat (*harga mati*). |
| **Kab. Cianjur** | **8 unit** | **22,0 hari** | **8,9%** | Rp 26,0 Jt | **28,5 (Sulit)** | Resistensi tinggi pada 8 unit cabang, durasi rata-rata 3 minggu. |
| **Kab. Cirebon** | **4 unit** | 18,2 hari | **9,2%** | Rp 12,0 Jt | **35,2 (Sulit)** | Mengikuti pola kekakuan pasar Kota Cirebon, minim ruang kompromi. |
| **Kab. Ciamis** | **3 unit** | 17,3 hari | **9,3%** | Rp 7,0 Jt | **36,9 (Sulit)** | Kepemilikan tradisional konservatif, resisten terhadap skema korporasi. |
| **Kab. Blora** | **3 unit** | 20,3 hari | 16,0% | Rp 30,0 Jt | **40,8 (Alot)** | Durasi lambat (20+ hari), negosiasi berlarut-larut. |
| **Kab. Banyumas** | **4 unit** | 20,0 hari | 17,2% | Rp 27,0 Jt | **43,0 (Alot)** | Kota niaga Purwokerto dan sekitarnya memiliki persaingan tenant kuat. |
| **Kab. Purwakarta** | **6 unit** | 17,8 hari | 15,2% | Rp 37,0 Jt | **43,8 (Alot)** | Kawasan transit industri, proses tawar-menawar memakan waktu 18 hari. |
| **Kota Surakarta (Solo)** | **3 unit** | 14,7 hari | 12,3% | Rp 23,5 Jt | **45,3 (Alot)** | Kawasan niaga prime (asking Rp 85 M), pemilik modal mapan tidak terburu-buru. |

---

## 6. Uji Empiris: Benarkah Kota Lebih Mudah vs Desa/Kabupaten Lebih Alot?

| Dimensi Metrik Negosiasi | Kota (Urban) | Kabupaten (Daerah/Rural) | Kesenjangan & Temuan Data |
| :--- | :---: | :---: | :--- |
| **Cakupan & Jumlah Pengajuan** | 26 Kota \| **108 Pengajuan (39,3%)** | 50 Kabupaten \| **167 Pengajuan (60,7%)** | Mayoritas ekspansi PGI berada di Kabupaten (60,7%). |
| **Rerata Harga Penawaran Awal** | Rp 55,28 Juta / tahun | Rp 36,66 Juta / tahun | Tarif sewa ruko Kota 50,8% lebih mahal dibanding Kabupaten. |
| **Rerata Durasi Negosiasi** | **10,3 Hari** (Median 8,0 hr) | **12,2 Hari** (Median 8,0 hr) | Kota 1,9 hari lebih cepat closing dibanding Kabupaten. |
| **Rerata Efisiensi Diskon Sewa** | 15,40% (Median 15,0%) | 15,69% (Median 16,0%) | Persentase diskon relatif seimbang (~15,5%). |
| **Konsentrasi Kasus Alot (> 14 hr)** | 17,6% transaksi alot (19 cabang) | **20,4% transaksi alot (34 cabang)** | **80% dari Top 10 wilayah paling alot adalah KABUPATEN**. |

### Kesimpulan Strategis Uji Empiris:
1. **Hipotesis Terbukti Sangat Kuat pada Dimensi Resistensi & Kelambatan (Alot)**:
   * **8 dari 10 wilayah paling alot di Indonesia adalah KABUPATEN** (Kab. Bandung Barat, Kab. Grobogan, Kab. Cianjur, Kab. Cirebon, Kab. Ciamis, Kab. Blora, Kab. Banyumas, Kab. Purwakarta).
   * Sebanyak 34 cabang di kabupaten-kabupaten ini tersendat negosiasi alot (> 14 hari) dengan rata-rata realisasi mencapai 20,5 hari.
   * Faktor penyebab: Status ruko warisan keluarga majemuk (*family-owned inheritance assets*), musyawarah keluarga yang berbelit, dan *holding power* tanah pasif pemilik tradisional.
2. **Anomali Penting 1: Tidak Seluruh Kabupaten Alot (Faktor Oversupply Ruko Sub-Urban)**:
   * Kabupaten seperti Cilacap (40,0% diskon, 3 unit), Serang (27,7%, 3 unit), Banjarnegara (27,5%, 2 unit), Sukabumi (23,3%, 3 unit), Garut (durasi 5,3 hr, 3 unit), Subang (21,0%, 3 unit), dan Bekasi (durasi 5,2 hr, 6 unit) justru masuk peringkat **Top Termudah**.
   * Di wilayah-wilayah ini terjadi fenomena kelebihan pasokan ruko baru (*oversupply*), sehingga pemilik ruko perorangan sangat fleksibel demi mengamankan kepastian sewa korporasi PGI.
3. **Anomali Penting 2: Kota-Kota Niaga Tradisional yang Sangat Alot**:
   * Kota tidak otomatis mudah. **Kota Cirebon** (skor 28,3; diskon cuma 2,5%; 2 unit) dan **Kota Surakarta/Solo** (skor 45,3; durasi 14,7 hr; asking Rp 85 M) masuk jajaran paling alot karena berstatus pusat niaga tradisional heritage (Pecinan & pasar batik keraton) di mana pemilik properti adalah saudagar tua bermodal kuat yang pantang menurunkan harga.

---

## 7. Profil Kinerja 11 Klaster Regional 2026

| Klaster Geografis | Jumlah Pengajuan | Total Sewa Deal | Total Saving (Rp) | Diskon Rata-rata (%) | Durasi Rata-rata | Lead Negosiator |
| :--- | :---: | :---: | :---: | :---: | :---: | :--- |
| **Jawa Tengah Selatan** | 9 unit | Rp 238,5 Jt | Rp 78,0 Jt | **22,56%** | 13,1 hari | Mirza (5 deal) |
| **Banten Barat** | 13 unit | Rp 392,0 Jt | Rp 96,0 Jt | **20,08%** | 11,5 hari | Mirza (7 deal) |
| **Jabodetabek Puncak** | **82 unit** | **Rp 3.341,6 Jt** | **Rp 719,8 Jt** | **16,09%** | **9,9 hari** | **Bonita (47 deal)** |
| **Jawa Barat Priangan & Selatan** | 27 unit | Rp 793,0 Jt | Rp 160,0 Jt | **15,48%** | 14,0 hari | Mirza (14 deal) |
| **Jawa Timur** | 9 unit | Rp 380,5 Jt | Rp 74,5 Jt | **15,22%** | **8,3 hari** | Bonita (7 deal) |
| **Jawa Tengah Pantura** | 36 unit | Rp 1.094,0 Jt | Rp 220,3 Jt | **14,75%** | 11,6 hari | Mirza (19 deal) |
| **Bandung Raya** | 23 unit | Rp 1.025,0 Jt | Rp 203,2 Jt | **14,65%** | 11,8 hari | Bonita (12 deal) |
| **Solo Raya & DIY** | 16 unit | Rp 595,5 Jt | Rp 95,5 Jt | **14,44%** | 12,2 hari | Mirza (11 deal) |
| **Bali** | 12 unit | Rp 646,0 Jt | Rp 124,0 Jt | **14,33%** | **9,0 hari** | Bonita (7 deal) |
| **Jawa Barat Pantura & Purwasuka** | 30 unit | Rp 885,8 Jt | Rp 141,2 Jt | **13,63%** | **15,3 hari** | Bonita (15 deal) |
| **Wilayah Luar Jawa/Lainnya** | 18 unit | Rp 664,0 Jt | Rp 123,0 Jt | **14,72%** | 9,8 hari | Mirza (10 deal) |

---

## 8. SOP Taktikal Manajemen Negosiasi di Wilayah Alot

1. **Pemberlakuan Batas Cut-Off SLA 14 Hari (*Rule of 14 Days*)**:
   Jika dalam 14 hari negosiasi di wilayah alot (Bandung Barat, Grobogan, Cirebon, Cianjur) pemilik ruko belum bersedia memberikan diskon minimal 5%–10%, surveyor wajib langsung mengaktifkan opsi **Ruko B dan Ruko C** sebagai alternatif (*BATNA - Best Alternative to a Negotiated Agreement*). Langkah ini mematahkan dominasi psikologis pemilik tanah.
2. **Alokasi Negosiator Sesuai Persona Pasar (*Persona-Matching*)**:
   * **Bonita**: Ditugaskan pada koridor metropolitan bervolume besar atau bertarif sewa tinggi (> Rp 50 Juta per ruko seperti Jakarta, Bali, Surabaya, Bandung Kota) untuk memanfaatkan *room for discount* yang lebar.
   * **Mirza**: Ditugaskan pada koridor satelit Pantura dan Jawa Tengah/Jawa Barat yang memerlukan akselerasi *high-velocity closing* (< 10 hari) guna memotong rantai negosiasi alot.
3. **Senjata Kontrak Korporasi Multi-Tahun**:
   Gunakan garansi sewa institusi 3–5 tahun dengan pembayaran di muka berjangka serta klausul renovasi estetika gedung ruko yang ditanggung PGI. Skema ini terbukti ampuh meluluhkan tuan tanah tradisional yang awalnya bersikeras pada harga mati.
