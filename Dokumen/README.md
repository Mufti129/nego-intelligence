# DIREKTORI DOKUMEN & LAPORAN ANALISIS UPC PUSAT GADAI INDONESIA (PGI)

Folder ini mengumpulkan seluruh dokumen laporan resmi, master dokumentasi analitik, dan panduan operasional Visual Studio Code (VSCode) untuk proyek analisis data Unit Pembukaan Cabang (UPC) 2024–2026.

---

## 1. DAFTAR BERKAS DOKUMEN

| No | Nama Berkas | Format | Kategori | Keterangan & Deskripsi |
| :---: | :--- | :---: | :---: | :--- |
| 1 | [**`Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/Dokumen/Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx) | Word (.docx) | **Laporan Resmi Direksi** | Laporan eksekutif lengkap formal (4,4 MB) memuat Bab 1 s/d 7, 9 diagram grafik 300 DPI, tabel ringkasan KPI, evaluasi surveyor, negosiasi Bonita vs Mirza, peta geospasial 76 wilayah, Vendor Risk Quadrant 2025 vs 2026, analisis wilayah Bali, serta lembar pengesahan. |
| 2 | [**`DOKUMENTASI_LENGKAP_PROSES_ANALISIS_UPC_PGI.md`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/Dokumen/DOKUMENTASI_LENGKAP_PROSES_ANALISIS_UPC_PGI.md) | Markdown (.md) | **Master Dokumentasi** | Dokumentasi teknis komprehensif *end-to-end* yang merinci latar belakang, arsitektur pembersihan data (1.047 $\rightarrow$ 1.028 baris unik), metodologi dekomposisi lead time, uji statistik Mann-Whitney U, uji korelasi Pearson/Spearman, formula indeks spasial, kuadran risiko vendor, dan panduan eksekusi VSCode. |
| 3 | [**`panduan_analisis_negosiasi_bulanan_vscode.md`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/Dokumen/panduan_analisis_negosiasi_bulanan_vscode.md) | Markdown (.md) | **Panduan VSCode** | Panduan operasional menjalankan skrip analisis performa negosiasi bulanan, komparasi Bonita vs Mirza, serta uji korelasi durasi vs efisiensi diskon (`analisa_performa_negosiasi_bulanan_2026.py`). |
| 4 | [**`panduan_analisis_geospasial_vscode.md`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/Dokumen/panduan_analisis_geospasial_vscode.md) | Markdown (.md) | **Panduan VSCode** | Panduan operasional pemetaan spasial indeks kemudahan negosiasi wilayah berbasis berkas TopoJSON 524 kabupaten/kota BPS (`analisa_geospasial_kemudahan_nego.py`). |
| 5 | [**`panduan_analisis_risiko_vendor_vscode.md`**](file:///Users/macbookair/Documents/Analisa/Nego_baru/Dokumen/panduan_analisis_risiko_vendor_vscode.md) | Markdown (.md) | **Panduan VSCode** | Panduan operasional pembuatan Vendor Risk Quadrant (2025 vs 2026), pelacakan vektor migrasi kontraktor, dan analisis dekomposisi kendala wilayah Bali pada Trimo (`analisa_risiko_vendor.py`). |

---

## 2. CARA PENGGUNAAN DOKUMEN

* **Untuk Keperluan Presentasi Direksi & Rapat Manajemen**:
  Buka berkas **`Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx`** menggunakan Microsoft Word. Dokumen ini sudah dilengkapi format tata letak elegan, palet korporat resmi, grafik beresolusi tinggi (300 DPI), serta lembar pengesahan tanda tangan.
* **Untuk Keperluan Dokumentasi Teknis & Pemahaman Metodologi**:
  Buka berkas **`DOKUMENTASI_LENGKAP_PROSES_ANALISIS_UPC_PGI.md`** di VSCode (klik kanan -> *Open Preview* atau `Cmd + Shift + V` di Mac / `Ctrl + Shift + V` di Windows) untuk membaca seluruh formula, logika filter data, uji hipotesis, dan alur diagram Mermaid.
* **Untuk Menjalankan Skrip Python Mandiri**:
  Gunakan panduan operasional masing-masing modul (`panduan_analisis_..._vscode.md`) untuk menjalankan skrip di terminal VSCode.
