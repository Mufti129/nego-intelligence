import os
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import pandas as pd
import numpy as np

def create_report():
    print("Reading cleaned dataset from data_nego_baru/...")
    f_xlsx = os.path.join("data_nego_baru", "data_nego_baru_cleaned_2024-2026.xlsx")
    if os.path.exists(f_xlsx):
        df = pd.read_excel(f_xlsx)
    else:
        df = pd.read_csv("properties_cleaned_2024-2026.csv")
    
    doc = Document()
    
    # Configure 1-inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
        # Configure Header & Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "PT PUSAT GADAI INDONESIA | DIVISI BISNIS — BUSINESS DEVELOPMENT"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.runs[0].font.size = Pt(8.5)
        hp.runs[0].font.color.rgb = RGBColor(120, 120, 120)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = "Laporan Analisis Komprehensif UPC PGI v3 (2024–2026) — Dokumen Konfidensial Internal"
        fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fp.runs[0].font.size = Pt(8.5)
        fp.runs[0].font.color.rgb = RGBColor(120, 120, 120)

    # Color Palette
    HEX_PRIMARY = "0F2C59"      # Deep Navy
    HEX_SECONDARY = "1E40AF"    # Royal Blue
    HEX_ACCENT = "0D9488"       # Teal
    HEX_LIGHT_BG = "F8FAFC"     # Slate 50
    HEX_BORDER = "CBD5E1"       # Slate 300
    HEX_DARK = "0F172A"         # Slate 900
    HEX_MUTED = "64748B"        # Slate 500
    
    COLOR_PRIMARY = RGBColor(15, 44, 89)
    COLOR_SECONDARY = RGBColor(30, 64, 175)
    COLOR_DARK = RGBColor(15, 23, 42)
    COLOR_MUTED = RGBColor(100, 116, 139)

    def set_cell_background(cell, fill_hex):
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading_elm)

    def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'''
            <w:tcMar {nsdecls("w")}>
                <w:top w:w="{top}" w:type="dxa"/>
                <w:bottom w:w="{bottom}" w:type="dxa"/>
                <w:left w:w="{left}" w:type="dxa"/>
                <w:right w:w="{right}" w:type="dxa"/>
            </w:tcMar>
        ''')
        tcPr.append(tcMar)

    def set_table_borders(table, color="D1D5DB", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:left w:val="none"/>
                <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:right w:val="none"/>
                <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="none"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    def add_callout(text, title="CATATAN PENTING", bg_hex="EFF6FF", border_hex="2563EB"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        set_cell_background(cell, bg_hex)
        set_cell_margins(cell, top=160, bottom=160, left=240, right=200)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)
        
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(2)
        cp.paragraph_format.space_after = Pt(2)
        run_title = cp.add_run(f"📌 {title}\n")
        run_title.font.name = "Arial"
        run_title.font.size = Pt(10)
        run_title.font.bold = True
        run_title.font.color.rgb = COLOR_PRIMARY
        
        run_text = cp.add_run(text)
        run_text.font.name = "Arial"
        run_text.font.size = Pt(9.5)
        run_text.font.color.rgb = COLOR_DARK
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_heading_3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(10.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_DARK
        return p

    def add_body(text, bold_prefix="", italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = "Arial"
            r_pre.font.size = Pt(10)
            r_pre.font.bold = True
            r_pre.font.color.rgb = COLOR_DARK
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(10)
        r.font.italic = italic
        r.font.color.rgb = COLOR_DARK
        return p

    def format_table(table, col_widths=None):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table)
        
        # Header Row
        hdr_cells = table.rows[0].cells
        for i, cell in enumerate(hdr_cells):
            set_cell_background(cell, HEX_PRIMARY)
            set_cell_margins(cell, top=140, bottom=140, left=160, right=160)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(9.5)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(255, 255, 255)
        
        # Body Rows
        for r_idx, row in enumerate(table.rows[1:]):
            bg = HEX_LIGHT_BG if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, cell in enumerate(row.cells):
                set_cell_background(cell, bg)
                set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(9)
                        r.font.color.rgb = COLOR_DARK
                        
        if col_widths:
            for row in table.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = Inches(width)

    def add_chart_image(image_path, width_inches=6.0, caption=""):
        if os.path.exists(image_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run()
            run.add_picture(image_path, width=Inches(width_inches))
            
            if caption:
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(2)
                p_cap.paragraph_format.space_after = Pt(8)
                r_cap = p_cap.add_run(f"📈 {caption}")
                r_cap.font.name = "Arial"
                r_cap.font.size = Pt(8.5)
                r_cap.font.italic = True
                r_cap.font.color.rgb = COLOR_MUTED

    print("Building Document Content...")

    # ==========================================
    # COVER PAGE
    # ==========================================
    p_org = doc.add_paragraph()
    p_org.paragraph_format.space_before = Pt(36)
    p_org.paragraph_format.space_after = Pt(4)
    r_org = p_org.add_run("PT PUSAT GADAI INDONESIA")
    r_org.font.name = "Arial"
    r_org.font.size = Pt(14)
    r_org.font.bold = True
    r_org.font.color.rgb = COLOR_PRIMARY

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_after = Pt(40)
    r_div = p_div.add_run("DIVISI BISNIS — BUSINESS DEVELOPMENT")
    r_div.font.name = "Arial"
    r_div.font.size = Pt(11)
    r_div.font.bold = True
    r_div.font.color.rgb = COLOR_SECONDARY

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run("LAPORAN ANALISIS KOMPREHENSIF\nPROSES PEMBUKAAN CABANG (UPC) — VERSI 3 (UPDATE DATA TERBARU)")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(21)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_PRIMARY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(46)
    r_sub = p_sub.add_run("Audit Lengkap Berkas 'properties (1) - properties (2024-2026)-.csv', Integrasi TopoJSON 524 Wilayah, Harmonisasi Sheet Diskon Negosiasi 2024–2026, dan Rapor Kontraktor")
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(11.5)
    r_sub.font.color.rgb = COLOR_MUTED

    # Metadata Card
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(meta_table, color="CBD5E1")
    
    meta_data = [
        ("Klasifikasi Dokumen", "KONFIDENSIAL — HANYA KALANGAN INTERNAL DIREKSI & MANAGEMENT"),
        ("Berkas Sumber Data", "properties (1) - properties (2024-2026)-.csv (1.047 baris mentah)"),
        ("Basis Analisis Bersih", "1.028 Cabang Unik (Dideduplikasi dari 19 histori termin pembayaran)"),
        ("Cakupan Geografis", "90 Kabupaten/Kota Terintegrasi TopoJSON Nasional (524 Wilayah BPS)"),
        ("Penyusun Dokumen", "Mukhammad Rekza Mufti (Data Analyst – Divisi Bisnis)"),
        ("Tanggal Pembaruan", "September 2026 (Pembaruan Menyeluruh Data Sheet 2024–2026)")
    ]
    for idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[idx]
        set_cell_background(row.cells[0], "F1F5F9")
        set_cell_background(row.cells[1], "FFFFFF")
        set_cell_margins(row.cells[0], top=80, bottom=80, left=120, right=120)
        set_cell_margins(row.cells[1], top=80, bottom=80, left=120, right=120)
        
        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.font.name = "Arial"
        r0.font.size = Pt(9.5)
        r0.font.bold = True
        r0.font.color.rgb = COLOR_PRIMARY
        
        p1 = row.cells[1].paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.name = "Arial"
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_DARK
        if idx == 0:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(185, 28, 28)

    doc.add_page_break()

    # ==========================================
    # RINGKASAN EKSEKUTIF
    # ==========================================
    add_heading_1("RINGKASAN EKSEKUTIF")
    add_body(
        "Laporan Versi 3 ini menyajikan pembaruan komprehensif atas proses pengadaan dan pembukaan Unit Pelayanan Cabang (UPC) "
        "Pusat Gadai Indonesia (PGI) untuk periode Januari 2024 hingga September 2026 berdasarkan dataset mutakhir "
        "dan lembar kerja resmi Laporan Negosiator. Seluruh kolom durasi 7 siklus dan finansial telah dirampingkan dan distandarisasi "
        "secara presisi. Melalui deduplikasi atas 19 baris multi-termin pembayaran dan pengeluaran 18 cabang bertanda (Pelebaran) (bukan unit baru UPC), "
        "analisis ini mengolah tepat 1.039 cabang ruko mandiri yang terintegrasi dengan pemetaan geospasial 92 kabupaten/kota di seluruh Indonesia."
    )

    # KPI Dashboard Table
    kpi_tbl = doc.add_table(rows=7, cols=3)
    format_table(kpi_tbl, [2.2, 2.2, 2.1])
    kpi_headers = ["Indikator Strategis", "Realisasi Kunci", "Catatan & Evaluasi Bisnis"]
    for i, h in enumerate(kpi_headers):
        kpi_tbl.rows[0].cells[i].paragraphs[0].text = h
        
    kpi_rows = [
        ("Total Cabang & Skala Anggaran", "1.039 Cabang Unik\nRp 109,2 M (Sewa Transfer)\nRp 127,4 M (Renovasi)", "Ekspansi masif di 92 Kab/Kota; komitmen sewa rata-rata 3–5 tahun per lokasi (eksklusif 18 unit pelebaran)."),
        ("Total Penghematan Riil (2024–2026)", "Rp 2,24 Miliar (Rp 2.243.333.335)\n• 2024: Rp 0 (0,0% diskon)\n• 2025: Rp 111,0 Jt (0,65% diskon)\n• 2026: Rp 2,13 Miliar (15,60% diskon)", "Evolusi efisiensi finansial sewa ruko: melonjak dari Rp 0 di 2024 menjadi Rp 2,13 Miliar di 2026."),
        ("Durasi Negosiasi Riil (durasi_nego_hari)", "Rerata Keseluruhan: 7,6 Hari\n• 2024: 1,4 hr (median 1 hr)\n• 2025: 9,4 hr (median 1 hr)\n• 2026: 11,5 hr (median 5 hr)", "Akselerasi durasi tawar-menawar riil di tahun 2026 ditopang oleh spesialisasi tim Mirza dan Bonita."),
        ("Kepatuhan SLA Renovasi Kontraktor", "Tren Menurun Konsisten:\n2024: 68,1% → 2025: 64,2%\n→ 2026: 53,8% (Krisis Kepatuhan)", "Kepatuhan dihitung per transaksi individual. Hampir separuh proyek di 2026 terlambat dari target operasional."),
        ("Waktu Tunggu Grand Opening (Tunggu GO)", "Rerata Keseluruhan: 9,7 Hari (Median 5,0 hr)\n• 2024: 14,0 hr (median 7,0 hr)\n• 2025: 9,3 hr (median 5,0 hr)\n• 2026: 5,5 hr (median 4,0 hr)", "Akselerasi persiapan akhir pasca renovasi (tgl_realisasi s/d Open_cabang) terpangkas 60% dari 14,0 hari ke 5,5 hari."),
        ("Lead Time Siklus Buka Cabang (UPC)", "Tahun 2026:\n• Pasca-Approval (Tahap 2–7): 58,6 Hari (~58 hr, med 53 hr)\n• Total Lead Time E2E (Tahap 1–7): 69,1 Hari (med 64 hr)\nRerata 3 Tahun Penuh: 69,0 Hari (med 60 hr)", "Harmonisasi definisi: Pasca-approval (approved s/d buka) konsisten 58,6 hari. Siklus penuh (pengajuan s/d buka) menyerap 69,1 hari (inklusif 10,7 hari pengajuan s/d approval).")
    ]
    for idx, (c1, c2, c3) in enumerate(kpi_rows):
        row = kpi_tbl.rows[idx+1]
        row.cells[0].paragraphs[0].text = c1
        row.cells[1].paragraphs[0].text = c2
        row.cells[2].paragraphs[0].text = c3

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_callout(
        "1. KOREKSI DISKON NEGOSIASI 2024–2026: Data sheet resmi (harga_awal_penawaran vs harga_rental_final) membuktikan "
        "bahwa tahun 2024 berjalan tanpa diskon (100% cabang deal di harga penawaran awal), 2025 mulai membukukan diskon Rp 111 Juta, "
        "dan puncaknya di 2026 menghasilkan penghematan fantastis sebesar Rp 2,13 Miliar (rata-rata diskon 15,60%).\n"
        "2. ANOMALI KEPATUHAN KONTRAKTOR: Kepatuhan SLA kontraktor mengalami pemburukan bertahap dari 68,1% menjadi 53,8%. "
        "Kontraktor Trimo memegang 185 proyek (volume terbesar ke-3) namun memiliki kepatuhan terburuk ke-2 (hanya 34,6%), "
        "terdistorsi beban penugasan khusus wilayah Bali (12 cabang kepatuhan 0%).",
        title="Dua Temuan Kritis Direksi"
    )

    # ==========================================
    # BAB 1: METODOLOGI & AUDIT DATA
    # ==========================================
    add_heading_1("BAB 1: LATAR BELAKANG, METODOLOGI & AUDIT KUALITAS DATA")
    
    add_heading_2("1.1 Konteks & Definisi Siklus 6 Tahap UPC")
    add_body(
        "Proses pembukaan cabang Unit Pelayanan Cabang (UPC) Pusat Gadai Indonesia mencakup 6 tahapan berurutan: "
        "pengajuan dokumen lokasi, persetujuan (approval) manajemen, negosiasi harga sewa ruko, renovasi fisik bangunan, "
        "persiapan sarana pendukung, hingga peresmian operasional (grand opening). Laporan ini mengevaluasi lead time, "
        "efisiensi biaya, serta performa personalia dan mitra eksternal di setiap etape tersebut."
    )

    t1_tbl = doc.add_table(rows=7, cols=3)
    format_table(t1_tbl, [1.8, 2.5, 2.2])
    t1_headers = ["Tahap Proses", "Definisi Operasional", "Parameter Database Terkait"]
    for i, h in enumerate(t1_headers):
        t1_tbl.rows[0].cells[i].paragraphs[0].text = h

    t1_data = [
        ("1. Pengajuan → Approved", "Durasi pembuatan berkas survei hingga terbit persetujuan formal manajemen.", "p.application_date s/d p.approved_at"),
        ("2. Tunggu Nego", "Jeda waktu tunggu pasca persetujuan hingga proses negosiasi pertama kali dibuka.", "p.approved_at s/d tanggal_awal_nego"),
        ("3. Realisasi Nego", "Total waktu negosiasi sewa dari hari pertama hingga kesepakatan deal tercapai.", "Kolom resmi: lama_waktu_realisasi_nego (hari)"),
        ("4. Tunggu Renovasi", "Masa jeda administrasi sewa hingga kontraktor mulai melakukan mobilisasi fisik.", "tgl_nego_berakhir s/d tgl_awal_renovasi"),
        ("5. Realisasi Renovasi", "Durasi pengerjaan fisik ruko oleh vendor kontraktor hingga serah terima lunas.", "tgl_awal_renovasi s/d paid_at / tgl_realisasi"),
        ("6. Tunggu Grand Opening", "Persiapan akhir (loading barang, instalasi brankas & IT, briefing karyawan).", "paid_at s/d Open_cabang (HRIS)")
    ]
    for idx, (c1, c2, c3) in enumerate(t1_data):
        row = t1_tbl.rows[idx+1]
        row.cells[0].paragraphs[0].text = c1
        row.cells[1].paragraphs[0].text = c2
        row.cells[2].paragraphs[0].text = c3

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_heading_2("1.2 Audit Kualitas Data & Rekayasa Perbaikan v3")
    add_body(
        "Pada versi 3 ini, seluruh kelemahan pencatatan data yang diidentifikasi pada laporan-laporan sebelumnya telah "
        "dibersihkan dan distandarisasi secara sistematis berdasarkan berkas mutakhir 'properties (1) - properties (2024-2026)-.csv':"
    )

    t2_tbl = doc.add_table(rows=6, cols=3)
    format_table(t2_tbl, [1.8, 2.3, 2.4])
    t2_headers = ["Isu Integritas Data", "Kondisi Awal / Cakupan", "Solusi & Resolusi Laporan v3"]
    for i, h in enumerate(t2_headers):
        t2_tbl.rows[0].cells[i].paragraphs[0].text = h

    t2_data = [
        ("Duplikasi Multi-Termin", "17 cabang memiliki 36 baris duplikasi karena multi-termin invoice pembayaran renovasi.", "Deduplikasi berbasis nomor_pengajuan dengan mengambil baris realisasi final (1.028 baris unik)."),
        ("Pencatatan Diskon Nego", "Berkas baru memuat Hargaawal_nego_sheet & hargarental_sheet lengkap 100% untuk 2024–2026.", "Aturan khusus: 2024–2025 dihitung berbasis nama_negosiator; 2026 berbasis nama_negosiator_2 & data sheet."),
        ("Field Approved_at Kosong", "Kosong pada 100% data 2024 dan 69,2% data 2025 (fitur approval baru aktif akhir 2025).", "Durasi approval 2024 ditandai N/A; kohort 2026 dianalisis penuh 100% (275 cabang terisi lengkap)."),
        ("Sentinel Date '1 Jan 2000'", "13 baris memiliki tanggal buka '1 January 2000' sebagai placeholder toko belum buka.", "Nilai dummy dinormalisasi menjadi status 'Belum Open' dan diexclude dari durasi pembukaan."),
        ("Ketiadaan Poligon Wilayah", "TopoJSON lama kehilangan 4 wilayah pemekaran (Bandung Barat, Pangandaran, Serang, Tangsel).", "Mengintegrasikan indonesia-kabkot-topo.json (524 wilayah BPS) mencakup 100% wilayah PGI.")
    ]
    for idx, (c1, c2, c3) in enumerate(t2_data):
        row = t2_tbl.rows[idx+1]
        row.cells[0].paragraphs[0].text = c1
        row.cells[1].paragraphs[0].text = c2
        row.cells[2].paragraphs[0].text = c3

    # ==========================================
    # BAB 2: KINERJA KESELURUHAN & SIKLUS WAKTU
    # ==========================================
    add_heading_1("BAB 2: KINERJA KESELURUHAN & ANALISIS BOTTLENECK SIKLUS")
    add_body(
        "Total durasi siklus pembukaan cabang dari pengajuan formulir hingga grand opening (Total Lead Time Penuh) "
        "berada pada rata-rata 68,8 hari (median 60,0 hari), sedangkan siklus operasional pasca-approval direksi "
        "(approved_at s/d Open_cabang) pada tahun 2026 tercatat sebesar 58,7 hari (~58 hari, median 53,0 hari). "
        "Meskipun total lead time relatif stabil dari tahun ke tahun, dekomposisi per tahapan memperlihatkan pergeseran "
        "yang sangat signifikan antara tahapan administratif negosiasi dan tahapan fisik operasional renovasi."
    )

    t3_tbl = doc.add_table(rows=5, cols=11)
    format_table(t3_tbl, [0.55, 0.45, 0.65, 0.55, 0.65, 0.55, 0.55, 0.65, 0.65, 0.90, 0.90])
    t3_headers = ["Tahun", "Unit", "1. Approval", "2. T. Nego", "3. Dur. Nego", "4. Berkas", "5. T. Renov", "6. Dur. Renov", "7. Tunggu GO*", "Pasca-Appr (2-7)", "Total Lead (1-7)"]
    for i, h in enumerate(t3_headers):
        t3_tbl.rows[0].cells[i].paragraphs[0].text = h

    t3_data = [
        ("2024", "317", "N/A*", "3,0 hr", "1,4 hr", "32,6 hr", "3,9 hr", "17,2 hr", "14,0 hr", "N/A*", "69,5 hr"),
        ("2025", "437", "13,5 hr*", "0,0 hr", "9,4 hr", "24,4 hr", "4,9 hr", "19,9 hr", "9,3 hr*", "57,2 hr", "68,5 hr"),
        ("2026", "285", "10,7 hr", "3,2 hr", "11,5 hr", "11,1 hr", "7,1 hr", "24,0 hr", "5,5 hr", "58,6 hr (~58h)", "69,1 hr"),
        ("Total", "1.039", "11,6 hr*", "1,9 hr", "7,6 hr", "23,2 hr", "5,2 hr", "20,2 hr", "9,7 hr", "58,2 hr*", "69,0 hr")
    ]
    for idx, r_data in enumerate(t3_data):
        row = t3_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    add_body(
        "Catatan Metrik & Harmonisasi Lead Time Pembukaan Cabang (7 Siklus Murni Kalender):\n"
        "• Pasca-Approval (Tahap 2–7): Berada pada rata-rata 58,6 hari (~58 hari di 2026, median 53,0 hari). Metrik ini mengukur durasi pelaksanaan operasional sejak cabang disetujui direksi (approved_at) hingga resmi buka (Open_cabang).\n"
        "• Total Lead Time Penuh (Tahap 1–7): Berada pada rata-rata 69,1 hari di 2026 (median 64,0 hari; rerata kumulatif 3 tahun 69,0 hari). Metrik ini mengukur siklus end-to-end menyeluruh sejak surveyor pertama kali mengajukan lokasi (application_date) hingga toko buka. Selisih antara jumlahan nominal 7 tahapan (72,8 hari) dengan lead time riil (69,1 hari) merefleksikan efisiensi fast-tracking (negosiasi paralel) sebesar 3,6 hari.\n"
        "• Penyelarasan Kalender: Seluruh metrik durasi kini dihitung berbasis hari kalender murni (.dt.normalize()), memastikan jeda operasional seperti TTD sewa ke awal renovasi tidak terdistorsi oleh jam pencatatan sistem.\n"
        "• (*) Pengajuan 2024 N/A karena approved_at belum diimplementasikan di sistem.",
        italic=True
    )

    add_chart_image("hasil_analisis/grafik/4_dekomposisi_siklus_waktu_upc.png", width_inches=6.0, caption="Grafik 2.1: Dekomposisi Lead Time Proses Pembukaan Cabang UPC (2024–2026)")

    add_heading_2("2.1 Analisis Titik Sumbatan (Bottleneck Analysis)")
    add_body(
        "Dua tahapan yang memakan porsi waktu terbesar dalam proses pembukaan cabang adalah Proses Negosiasi (rata-rata 23,6 hari berbasis kolom resmi lama_waktu_realisasi_nego) "
        "dan Proses Renovasi Fisik (rata-rata 20,1 hari). Gabungan kedua proses ini menyerap lebih dari 63% total lead time ekspansi.\n"
        "Pada tahun 2026, efisiensi negosiasi berhasil memangkas durasi tawar-menawar menjadi hanya 11,5 hari (turun 65% dari 2024). "
        "Namun, waktu pengerjaan renovasi kontraktor justru membengkak menjadi 23,9 hari (naik 40%), sehingga percepatan grand opening "
        "sebagian tertahan oleh kelambatan pengerjaan fisik ruko.\n"
        "Di sisi lain, terdapat perbaikan operasional yang sangat impresif pada tahap Waktu Tunggu Grand Opening (jeda dari renovasi selesai/paid_at hingga cabang resmi beroperasi). "
        "Waktu tunggu ini berhasil dipangkas drastis sebesar 60%, dari rata-rata 14,0 hari (2024) menjadi 9,4 hari (2025), dan kini hanya 5,6 hari (median 4,0 hari di 2026). "
        "Hal ini membuktikan efisiensi tinggi tim operasional cabang, IT, dan logistik dalam memobilisasi sarana toko begitu renovasi diserahterimakan."
    )

    # ==========================================
    # BAB 3: KINERJA SURVEYOR
    # ==========================================
    add_heading_1("BAB 3: KINERJA PENGAJUAN & SURVEYOR LOKASI")
    add_body(
        "Sebanyak 1.028 cabang diajukan oleh 9 personil surveyor. Terdapat konsentrasi beban kerja yang sangat ekstrem, "
        "di mana 3 surveyor teratas menangani 80,8% dari total keseluruhan pengajuan lokasi baru di Indonesia."
    )

    t4_tbl = doc.add_table(rows=10, cols=4)
    format_table(t4_tbl, [2.2, 1.4, 1.4, 1.5])
    t4_headers = ["Nama Surveyor", "Total Pengajuan", "Porsi Volume (%)", "Tahun Aktif"]
    for i, h in enumerate(t4_headers):
        t4_tbl.rows[0].cells[i].paragraphs[0].text = h

    t4_data = [
        ("Djadjang", "320", "31,1%", "2024, 2025, 2026"),
        ("Salma", "294", "28,6%", "2024, 2025, 2026"),
        ("Hanif", "217", "21,1%", "2024, 2025, 2026"),
        ("Dika", "58", "5,6%", "2025, 2026"),
        ("Dara", "58", "5,6%", "2026"),
        ("Sinta", "30", "2,9%", "2024, 2025"),
        ("Bonita", "24", "2,3%", "2025, 2026"),
        ("Gita", "20", "1,9%", "2024, 2025"),
        ("Claher Prastian", "7", "0,7%", "2024, 2025")
    ]
    for idx, r_data in enumerate(t4_data):
        row = t4_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    add_body(
        "Tingginya konsentrasi pada Djadjang, Salma, dan Hanif menunjukkan ketergantungan operasional yang tinggi pada trio surveyor ini. "
        "Pada tahun 2026, kemunculan Dara (58 pengajuan) memberikan diversifikasi positif bagi pemerataan beban survei lapangan."
    )

    # ==========================================
    # BAB 4: KINERJA NEGOSIASI 2024-2026
    # ==========================================
    add_heading_1("BAB 4: ANALISIS KINERJA NEGOSIASI & NEGOSIATOR (2024–2026)")
    
    add_heading_2("4.1 Evaluasi Negosiasi Historis (2024–2025) — Basis: nama_negosiator")
    add_body(
        "Berdasarkan pembaruan data sheet resmi, pada tahun 2024 seluruh deal (317 cabang) ditutup persis di harga penawaran awal "
        "(diskon 0,0%). Inisiatif penghematan mulai muncul pada akhir 2025 sebesar Rp 111,0 Juta di 15 cabang ruko. "
        "Evaluasi negosiator 2024–2025 menggunakan kolom nama_negosiator dan durasi resmi lama_waktu_realisasi_nego:"
    )

    t5_tbl = doc.add_table(rows=9, cols=6)
    format_table(t5_tbl, [1.6, 0.7, 1.2, 1.0, 1.0, 1.0])
    t5_headers = ["Negosiator", "Deal", "Total Asking", "Total Saving", "Avg Diskon", "Avg Durasi"]
    for i, h in enumerate(t5_headers):
        t5_tbl.rows[0].cells[i].paragraphs[0].text = h

    t5_data = [
        ("Sinta", "304", "Rp 9,61 Miliar", "Rp 0", "0,00%", "31,7 hari"),
        ("Salma", "179", "Rp 5,98 Miliar", "Rp 14,0 Jt", "0,27%", "34,5 hari"),
        ("Mirza", "151", "Rp 4,90 Miliar", "Rp 42,0 Jt", "0,95%", "8,3 hari"),
        ("Bonita", "72", "Rp 2,76 Miliar", "Rp 49,0 Jt", "0,90%", "34,9 hari"),
        ("Dika", "20", "Rp 634,0 Jt", "Rp 6,0 Jt", "1,00%", "30,7 hari"),
        ("Gita", "15", "Rp 479,0 Jt", "Rp 0", "0,00%", "36,7 hari"),
        ("Claher Prastian", "9", "Rp 307,0 Jt", "Rp 0", "0,00%", "33,0 hari"),
        ("Triani", "3", "Rp 91,0 Jt", "Rp 0", "0,00%", "17,3 hari")
    ]
    for idx, r_data in enumerate(t5_data):
        row = t5_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    add_heading_2("4.2 Evaluasi Negosiasi Riil 2026 — Tim Negosiasi Resmi (Bonita & Mirza)")
    add_body(
        "Tahun 2026 mencatatkan lompatan efisiensi finansial tertinggi sepanjang sejarah ekspansi PGI. "
        "Sesuai arahan dan struktur operasional resmi manajemen, Tim Negosiasi 2026 secara mandiri diemban oleh dua personil spesialis: "
        "Bonita dan Mirza (menangani 282 cabang atau 96,2% transaksi nasional). Sementara itu, 11 cabang lainnya merupakan pengajuan langsung surveyor "
        "(Dika 10 cabang, Salma 1 cabang) yang ditutup langsung di harga penawaran awal (diskon 0,0%, saving Rp 0) tanpa melalui proses negosiasi sewa.\n"
        "Evaluasi negosiasi menerapkan harga_awal_penawaran sebagai harga awal, harga_rental_final sebagai harga sewa final, "
        "efisiensi_diskon_pct sebagai tolok ukur diskon, serta durasi_nego_hari sebagai durasi tawar-menawar:"
    )

    t6_tbl = doc.add_table(rows=6, cols=7)
    format_table(t6_tbl, [1.6, 0.6, 1.1, 1.1, 1.1, 0.9, 0.8])
    t6_headers = ["Kategori / Personil", "Deal", "Total Asking", "Total Deal", "Total Saving", "Avg Diskon", "Avg Durasi"]
    for i, h in enumerate(t6_headers):
        t6_tbl.rows[0].cells[i].paragraphs[0].text = h

    t6_data = [
        ("Bonita (Tim Nego Resmi)", "142", "Rp 7,40 M", "Rp 6,00 M", "Rp 1,40 Miliar", "18,25%", "8,5 hr"),
        ("Mirza (Tim Nego Resmi)", "132", "Rp 4,97 M", "Rp 4,24 M", "Rp 728,4 Jt", "14,05%", "15,5 hr"),
        ("Subtotal Tim Nego Resmi", "274", "Rp 12,37 M", "Rp 10,24 M", "Rp 2,13 Miliar", "16,22%", "11,9 hr"),
        ("Surveyor (Dika & Salma)*", "11", "Rp 306 Jt", "Rp 306 Jt", "Rp 0", "0,00%", "1,0 hr"),
        ("TOTAL TAHUN 2026", "285", "Rp 12,68 M", "Rp 10,54 M", "Rp 2,13 Miliar", "15,60%", "11,5 hr")
    ]
    for idx, r_data in enumerate(t6_data):
        row = t6_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx in [2, 4]:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    add_body("Catatan: (*) Dika (10 cabang) dan Salma (1 cabang) merupakan personil tim survei lapangan, bukan tim negosiasi.", italic=True)

    add_heading_2("4.3 Uji Signifikansi Statistik: Bonita vs Mirza (Kohort 2026)")
    add_body(
        "Bonita dan Mirza mendominasi 100% dari seluruh capaian penghematan sewa ruko tahun 2026 (Rp 2,13 Miliar). Pengujian statistik "
        "non-parametrik Mann-Whitney U membuktikan bahwa efisiensi diskon antara Bonita (rata-rata 18,25%, median 19,23%) "
        "dan Mirza (rata-rata 14,05%, median 13,33%) memperlihatkan pola spesialisasi portofolio yang sangat jelas dan saling melengkapi."
    )
    add_body(
        "Karakteristik Spesialisasi:\n"
        "• Bonita (The High-Value & Corporate Specialist): Menangani ruko tier atas bernilai tinggi (total asking Rp 7,40 Miliar). "
        "Mencatatkan penghematan masif sebesar Rp 1,40 Miliar (65,8% total saving perusahaan) dengan tingkat keberhasilan diskon 95,8% "
        "dan kecepatan rata-rata 8,5 hari.\n"
        "• Mirza (The Speed & Volume Specialist): Menangani volume ekspansi masif di koridor sekunder dan regional (132 deal, total asking Rp 4,97 Miliar), "
        "menghasilkan penghematan Rp 728,4 Juta dengan tingkat keberhasilan diskon 93,2% dan rata-rata diskon 14,05%."
    )

    add_heading_2("4.4 Tren Performa Negosiasi Bulanan Tahun 2026 (Januari – Agustus)")
    add_body(
        "Dekomposisi performa negosiasi bulanan tahun 2026 (berdasarkan tanggal deal tgl_nego_berakhir) memperlihatkan "
        "dinamika ekspansi yang sangat agresif pada Kuartal 2 (April–Juni) serta efisiensi durasi yang meningkat pesat pada Kuartal 3:"
    )

    t6b_tbl = doc.add_table(rows=10, cols=8)
    format_table(t6b_tbl, [1.3, 0.5, 0.9, 0.9, 0.9, 0.7, 0.7, 0.6])
    t6b_headers = ["Bulan 2026", "Deal", "Total Asking", "Total Sewa", "Total Saving", "Diskon %", "Durasi hr", "Success"]
    for i, h in enumerate(t6b_headers):
        t6b_tbl.rows[0].cells[i].paragraphs[0].text = h

    t6b_data = [
        ("Januari 2026", "22", "Rp 712,0 Jt", "Rp 574,0 Jt", "Rp 138,0 Jt", "18,84%", "3,5 hr", "86,4%"),
        ("Februari 2026", "39", "Rp 1,59 M", "Rp 1,32 M", "Rp 261,1 Jt", "14,96%", "9,8 hr", "84,6%"),
        ("Maret 2026", "36", "Rp 1,42 M", "Rp 1,13 M", "Rp 282,0 Jt", "18,37%", "17,8 hr", "88,9%"),
        ("April 2026", "54", "Rp 2,37 M", "Rp 2,01 M", "Rp 365,2 Jt", "15,20%", "14,6 hr", "92,6%"),
        ("Mei 2026", "36", "Rp 1,86 M", "Rp 1,52 M", "Rp 332,5 Jt", "15,74%", "13,2 hr", "97,2%"),
        ("Juni 2026", "42", "Rp 1,88 M", "Rp 1,57 M", "Rp 312,7 Jt", "13,60%", "7,0 hr", "90,5%"),
        ("Juli 2026", "37", "Rp 1,94 M", "Rp 1,64 M", "Rp 298,6 Jt", "14,18%", "14,4 hr", "91,9%"),
        ("Agustus 2026", "19", "Rp 911,0 Jt", "Rp 768,7 Jt", "Rp 142,3 Jt", "15,96%", "4,2 hr", "94,7%"),
        ("TOTAL 2026", "285", "Rp 12,68 M", "Rp 10,54 M", "Rp 2,13 Miliar", "15,60%", "11,5 hr", "90,9%")
    ]
    for idx, r_data in enumerate(t6b_data):
        row = t6b_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx == len(t6b_data) - 1:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    add_body(
        "Temuan Analisis Tren Bulanan:\n"
        "1. Puncak Penghematan Finansial: Terjadi pada April (Rp 365,2 Juta), Mei (Rp 332,5 Juta), dan Juni (Rp 312,7 Juta). Dalam tiga bulan ini saja, "
        "PGI berhasil menghemat biaya sewa ruko lebih dari Rp 1,01 Miliar.\n"
        "2. Puncak Volume Ekspansi: Kuartal 2 menyerap 132 cabang (46,3% dari seluruh penambahan cabang tahun 2026), "
        "dengan rekor tertinggi di April (54 deal).\n"
        "3. Peningkatan Kecepatan Negosiasi: Memasuki Kuartal 3, durasi negosiasi terpangkas signifikan menjadi 7,0 hari (Juni) dan 4,2 hari (Agustus)."
    )

    add_chart_image("hasil_analisis/grafik/1_tren_volume_dan_saving_2026.png", width_inches=6.0, caption="Grafik 4.1: Tren Bulanan Volume Deal dan Akumulasi Penghematan Sewa (Saving) Tahun 2026")

    add_heading_2("4.5 Tren Bulanan 2026: Dinamika Volume Deal & Penghematan Sewa — Bonita")
    add_body(
        "Tabel berikut menyajikan rekapitulasi tren bulanan performa Bonita sepanjang Januari hingga Agustus 2026, mencakup "
        "volume deal, nilai penawaran awal (asking price), harga kesepakatan sewa (deal price), penghematan sewa (saving), "
        "persentase efisiensi diskon, rata-rata durasi penuntasan negosiasi, serta tingkat keberhasilan (success rate):"
    )

    t6_bonita_tbl = doc.add_table(rows=10, cols=8)
    format_table(t6_bonita_tbl, [1.3, 0.5, 0.9, 0.9, 0.9, 0.7, 0.7, 0.6])
    t6_bonita_headers = ["Bulan 2026", "Deal", "Total Asking", "Total Sewa", "Total Saving", "Diskon %", "Durasi hr", "Success"]
    for i, h in enumerate(t6_bonita_headers):
        t6_bonita_tbl.rows[0].cells[i].paragraphs[0].text = h

    t6_bonita_data = [
        ("Januari 2026", "13", "Rp 458,0 Jt", "Rp 347,0 Jt", "Rp 111,0 Jt", "23,29%", "1,0 hr", "100,0%"),
        ("Februari 2026", "16", "Rp 887,2 Jt", "Rp 712,0 Jt", "Rp 175,2 Jt", "18,10%", "11,8 hr", "100,0%"),
        ("Maret 2026", "19", "Rp 897,0 Jt", "Rp 700,0 Jt", "Rp 197,0 Jt", "21,25%", "15,7 hr", "94,7%"),
        ("April 2026", "26", "Rp 1,40 M", "Rp 1,18 M", "Rp 219,0 Jt", "16,68%", "11,7 hr", "92,3%"),
        ("Mei 2026", "17", "Rp 1,02 M", "Rp 808,0 Jt", "Rp 214,5 Jt", "18,38%", "10,6 hr", "94,1%"),
        ("Juni 2026", "23", "Rp 1,27 M", "Rp 1,03 M", "Rp 241,5 Jt", "16,40%", "5,9 hr", "95,7%"),
        ("Juli 2026", "18", "Rp 971,0 Jt", "Rp 814,8 Jt", "Rp 156,2 Jt", "15,63%", "3,8 hr", "94,4%"),
        ("Agustus 2026", "10", "Rp 495,0 Jt", "Rp 405,5 Jt", "Rp 89,5 Jt", "19,10%", "2,4 hr", "100,0%"),
        ("TOTAL 2026", "142", "Rp 7,40 M", "Rp 6,00 M", "Rp 1,40 Miliar", "18,25%", "8,5 hr", "95,8%")
    ]
    for idx, r_data in enumerate(t6_bonita_data):
        row = t6_bonita_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx == len(t6_bonita_data) - 1:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    add_body(
        "Insight Kunci Kinerja Bonita:\n"
        "• Rekor Penghematan Sewa: Bonita mencetak efisiensi biaya tertinggi di bulan Juni (Rp 241,5 Juta), April (Rp 219,0 Juta), dan Mei (Rp 214,5 Juta).\n"
        "• Stabilitas Diskon Tinggi: Rata-rata diskon bulanan Bonita secara konsisten melampaui 15,6% setiap bulannya (puncak di Januari 23,29% dan Maret 21,25%).\n"
        "• Akselerasi Kecepatan Closing: Di Kuartal 3, durasi negosiasi Bonita terpangkas tajam menjadi 5,9 hari (Juni), 3,8 hari (Juli), dan 2,4 hari (Agustus)."
    )

    add_heading_2("4.6 Tren Bulanan 2026: Dinamika Volume Deal & Penghematan Sewa — Mirza")
    add_body(
        "Tabel berikut menyajikan rekapitulasi tren bulanan performa Mirza sepanjang Januari hingga Agustus 2026, mencakup "
        "volume deal, nilai penawaran awal (asking price), harga kesepakatan sewa (deal price), penghematan sewa (saving), "
        "persentase efisiensi diskon, rata-rata durasi penuntasan negosiasi, serta tingkat keberhasilan (success rate):"
    )

    t6_mirza_tbl = doc.add_table(rows=10, cols=8)
    format_table(t6_mirza_tbl, [1.3, 0.5, 0.9, 0.9, 0.9, 0.7, 0.7, 0.6])
    t6_mirza_headers = ["Bulan 2026", "Deal", "Total Asking", "Total Sewa", "Total Saving", "Diskon %", "Durasi hr", "Success"]
    for i, h in enumerate(t6_mirza_headers):
        t6_mirza_tbl.rows[0].cells[i].paragraphs[0].text = h

    t6_mirza_data = [
        ("Januari 2026", "6", "Rp 152,0 Jt", "Rp 125,0 Jt", "Rp 27,0 Jt", "18,60%", "10,0 hr", "100,0%"),
        ("Februari 2026", "17", "Rp 534,5 Jt", "Rp 448,7 Jt", "Rp 85,8 Jt", "17,27%", "11,1 hr", "100,0%"),
        ("Maret 2026", "15", "Rp 475,0 Jt", "Rp 390,0 Jt", "Rp 85,0 Jt", "17,18%", "22,6 hr", "93,3%"),
        ("April 2026", "28", "Rp 972,0 Jt", "Rp 825,8 Jt", "Rp 146,2 Jt", "13,84%", "17,3 hr", "92,9%"),
        ("Mei 2026", "19", "Rp 835,0 Jt", "Rp 717,0 Jt", "Rp 118,0 Jt", "13,39%", "15,6 hr", "100,0%"),
        ("Juni 2026", "19", "Rp 614,2 Jt", "Rp 543,0 Jt", "Rp 71,2 Jt", "10,20%", "8,5 hr", "84,2%"),
        ("Juli 2026", "19", "Rp 971,0 Jt", "Rp 828,6 Jt", "Rp 142,4 Jt", "12,80%", "24,5 hr", "89,5%"),
        ("Agustus 2026", "9", "Rp 416,0 Jt", "Rp 363,2 Jt", "Rp 52,8 Jt", "12,47%", "6,2 hr", "88,9%"),
        ("TOTAL 2026", "132", "Rp 4,97 M", "Rp 4,24 M", "Rp 728,4 Jt", "14,05%", "15,5 hr", "93,2%")
    ]
    for idx, r_data in enumerate(t6_mirza_data):
        row = t6_mirza_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx == len(t6_mirza_data) - 1:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    add_body(
        "Insight Kunci Kinerja Mirza:\n"
        "• Pendorong Utama Volume Ekspansi: Mirza menyerap beban deal terbesar pada masa puncak ekspansi di bulan April (28 deal) dan Mei (19 deal).\n"
        "• Efisiensi Finansial Konsisten: Menghasilkan saving di atas Rp 140 Juta pada bulan April (Rp 146,2 Juta) dan Juli (Rp 142,4 Juta).\n"
        "• Fleksibilitas Teritori Sekunder: Mirza banyak diterjunkan ke area-area dengan karakter negosiasi alot dan geografis menantang, membukukan rata-rata diskon kanonikal 14,05%."
    )

    add_heading_2("4.7 Matriks Kinerja Negosiator Bulanan: Bonita vs Mirza (Tim Nego)")
    add_body(
        "Kompilasi kontribusi bulanan memperlihatkan perbandingan volume deal, efisiensi diskon, penghematan biaya sewa, "
        "dan kecepatan durasi antara dua negosiator resmi PGI (Bonita dan Mirza):"
    )

    t6c_tbl = doc.add_table(rows=10, cols=11)
    format_table(t6c_tbl, [0.75, 0.4, 0.6, 0.7, 0.5, 0.4, 0.6, 0.7, 0.5, 0.55, 0.8])
    t6c_headers = ["Bulan", "Unit B", "Disk B", "Saving B", "Dur B", "Unit M", "Disk M", "Saving M", "Dur M", "Non-Tim*", "Total Saving"]
    for i, h in enumerate(t6c_headers):
        t6c_tbl.rows[0].cells[i].paragraphs[0].text = h

    t6c_data = [
        ("Januari", "13", "23,3%", "Rp 111,0 Jt", "1,0 hr", "6", "18,6%", "Rp 27,0 Jt", "10,0 hr", "3", "Rp 138,0 Jt"),
        ("Februari", "16", "18,1%", "Rp 175,2 Jt", "11,8 hr", "17", "17,3%", "Rp 85,8 Jt", "11,1 hr", "6", "Rp 261,1 Jt"),
        ("Maret", "19", "21,3%", "Rp 197,0 Jt", "15,7 hr", "15", "17,2%", "Rp 85,0 Jt", "22,6 hr", "2", "Rp 282,0 Jt"),
        ("April", "26", "16,7%", "Rp 219,0 Jt", "11,7 hr", "28", "13,8%", "Rp 146,2 Jt", "17,3 hr", "0", "Rp 365,2 Jt"),
        ("Mei", "17", "18,4%", "Rp 214,5 Jt", "10,6 hr", "19", "13,4%", "Rp 118,0 Jt", "15,6 hr", "0", "Rp 332,5 Jt"),
        ("Juni", "23", "16,4%", "Rp 241,5 Jt", "5,9 hr", "19", "10,2%", "Rp 71,2 Jt", "8,5 hr", "0", "Rp 312,7 Jt"),
        ("Juli", "18", "15,6%", "Rp 156,2 Jt", "3,8 hr", "19", "12,8%", "Rp 142,4 Jt", "24,5 hr", "0", "Rp 298,6 Jt"),
        ("Agustus", "10", "19,1%", "Rp 89,5 Jt", "2,4 hr", "9", "12,5%", "Rp 52,8 Jt", "6,2 hr", "0", "Rp 142,3 Jt"),
        ("TOTAL", "142", "18,3%", "Rp 1,40 M", "8,5 hr", "132", "14,0%", "Rp 728,4 Jt", "15,5 hr", "11", "Rp 2,13 M")
    ]
    for idx, r_data in enumerate(t6c_data):
        row = t6c_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx == len(t6c_data) - 1:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    add_body("Catatan: (*) Kolom 'Non-Tim' adalah 11 cabang yang diajukan langsung oleh surveyor lapangan (Dika: 10 deal, Salma: 1 deal) dengan diskon 0% (deal di harga asking). Tim negosiasi resmi 2026 secara eksklusif diemban oleh Bonita dan Mirza.", italic=True)

    add_chart_image("hasil_analisis/grafik/2_komparasi_saving_bonita_mirza_2026.png", width_inches=6.0, caption="Grafik 4.2: Komparasi Penghematan Sewa Bulanan: Bonita vs Mirza (Tahun 2026)")
    add_chart_image("hasil_analisis/grafik/3_komparasi_diskon_dan_durasi_2026.png", width_inches=6.2, caption="Grafik 4.3: Perbandingan Efisiensi Diskon (%) dan Kecepatan Realisasi Negosiasi (Hari)")

    add_heading_2("4.8 Analisis Distribusi & Uji Korelasi: Durasi Negosiasi vs Efisiensi Diskon Sewa")
    add_body(
        "Salah satu pertanyaan strategis fundamental bagi manajemen ekspansi adalah: Apakah proses negosiasi yang memakan waktu lebih lama "
        "berkorelasi dengan perolehan efisiensi diskon yang lebih tinggi, ataukah justru memperpanjang siklus tanpa pertambahan diskon yang signifikan? "
        "Untuk menjawab pertanyaan tersebut secara ilmiah, dilakukan uji korelasi bivariat parametrik (Pearson r) dan non-parametrik (Spearman rho), "
        "analisis regresi linier, serta dekomposisi distribusi kelompok durasi terhadap seluruh 275 transaksi cabang tahun 2026."
    )

    add_chart_image("hasil_analisis/grafik/7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png", width_inches=6.4, caption="Grafik 4.4: Sebaran (Scatter Plot) & Distribusi Durasi Negosiasi vs Efisiensi Diskon Sewa (Tahun 2026)")

    t_dist_tbl = doc.add_table(rows=7, cols=7)
    format_table(t_dist_tbl, [1.6, 0.7, 0.9, 0.8, 0.9, 0.9, 0.7])
    t_dist_headers = ["Kelompok Durasi Negosiasi", "Jumlah Deal", "Rerata Diskon", "Median Diskon", "Rerata Saving", "Total Saving", "Sukses Diskon"]
    for i, h in enumerate(t_dist_headers):
        t_dist_tbl.rows[0].cells[i].paragraphs[0].text = h

    t_dist_data = [
        ("1–5 hari (Super Cepat)", "64 deal", "16,41%", "16,5%", "Rp 6,44 Jt", "Rp 412,0 Jt", "96,9%"),
        ("6–10 hari (Cepat / Ideal)", "126 deal", "15,28%", "15,5%", "Rp 7,21 Jt", "Rp 908,8 Jt", "92,9%"),
        ("11–14 hari (Batas Toleransi SLA)", "32 deal", "17,72%", "16,5%", "Rp 9,15 Jt", "Rp 292,7 Jt", "93,8%"),
        ("15–20 hari (Zona Alot)", "13 deal", "14,46%", "13,0%", "Rp 7,31 Jt", "Rp 95,0 Jt", "84,6%"),
        ("> 20 hari (Zona Sangat Alot)", "40 deal", "13,83%", "13,0%", "Rp 8,18 Jt", "Rp 327,0 Jt", "72,5%"),
        ("TOTAL TAHUN 2026", "275 deal", "15,57%", "15,0%", "Rp 7,40 Jt", "Rp 2.035,5 Jt", "90,5%")
    ]
    for idx, r_data in enumerate(t_dist_data):
        row = t_dist_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx == len(t_dist_data) - 1:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_body(
        "Temuan Kunci Uji Korelasi & Analisis Distribusi:\n"
        "1. Ketiadaan Korelasi Positif (No Positive Correlation):\n"
        "   Hasil uji statistik menunjukkan koefisien korelasi Pearson r = -0,1040 (p = 0,085) dan Spearman rho = -0,0844 (p = 0,163). Nilai R-squared tercatat hanya 0,0108, "
        "   yang membuktikan bahwa hanya 1,1% variasi diskon yang berhubungan dengan lamanya durasi negosiasi, sementara 98,9% lainnya ditentukan oleh faktor eksternal "
        "   (ruang tawar lokasi, kematangan negosiator, dan profil psikologis pemilik ruko). Arah korelasi yang bernilai sedikit negatif mengonfirmasi bahwa "
        "   bernegosiasi lebih lama TIDAK menghasilkan diskon yang lebih besar.\n"
        "2. Zona Emas Transaksi (1–10 Hari):\n"
        "   Mayoritas mutlak transaksi sukses (69,1% atau 190 dari 275 cabang) berhasil dituntaskan dalam waktu 1–10 hari kalender dengan tingkat keberhasilan meraih diskon mencapai 93%–97%. "
        "   Bahkan pada kelompok super cepat (1–5 hari), diskon rata-rata mencapai 16,41% dengan 96,9% transaksi berhasil memperoleh potongan harga sewa.\n"
        "3. Hukum Penurunan Hasil Negosiasi (Law of Diminishing Returns) & Zona Redflag (> 14 Hari):\n"
        "   Ketika proses negosiasi melampaui batas 14 hari, efisiensi diskon rata-rata justru menurun menjadi 14,46% (15–20 hari) dan anjlok ke 13,83% (> 20 hari). "
        "   Yang paling kritis, tingkat kegagalan meraih diskon meningkat drastis: pada kelompok > 20 hari, 27,5% cabang berakhir deal tanpa diskon sama sekali (0% diskon). "
        "   Durasi negosiasi yang panjang pada kelompok ini merupakan manifestasi dari alotnya resistensi tuan tanah tradisional, bukan karena negosiator sedang memperjuangkan tambahan margin.\n"
        "4. Uji Durasi Masa Sewa Kontrak (Tenor Ruko):\n"
        "   Apabila korelasi diuji terhadap durasi masa sewa ruko (lama tahun kontrak), data membuktikan bahwa 92,7% cabang (255 dari 275 cabang) telah distandarisasi "
        "   pada tenor komitmen 5 tahun (r = 0,0027, p = 0,964). Dengan demikian, keunggulan diskon PGI bukan berasal dari variasi panjangnya tahun sewa, melainkan dari "
        "   daya tawar institusi PGI dan kemahiran taktis negosiator dalam mengeksekusi penawaran tunai korporasi di awal."
    )

    # ==========================================
    # BAB 5: PETA GEOSPASIAL & ANALISIS KARAKTERISTIK KEMUDAHAN NEGOSIASI
    # ==========================================
    add_heading_1("BAB 5: PETA GEOSPASIAL & ANALISIS KARAKTERISTIK KEMUDAHAN NEGOSIASI WILAYAH")
    add_body(
        "Tingkat keberhasilan negosiasi properti ruko sangat dipengaruhi oleh karakteristik pasar properti lokal, struktur kepemilikan aset, "
        "serta latar belakang demografis dan sosiokultural pemilik properti. Untuk mengukur efektivitas proses tawar-menawar secara objektif, "
        "disusun Skor Kemudahan Negosiasi Wilayah (skala 0–100) yang menggabungkan dua indikator kunci berbobot seimbang (50:50):\n"
        "1. Skor Kecepatan Closing (50%): Diukur dari durasi tawar-menawar riil (lama_waktu_realisasi_nego). Semakin singkat hari realisasi, semakin tinggi skor kemudahan.\n"
        "2. Skor Efisiensi Diskon (50%): Diukur dari persentase pemotongan harga sewa (efisiensi_diskon_pct). Semakin besar diskon yang diraih, semakin tinggi skor kemudahan.\n"
        "Hasil pengolahan terhadap 76 kabupaten/kota aktif pada tahun 2026 dipetakan secara spasial menggunakan berkas TopoJSON 524 wilayah BPS (indonesia-kabkot-topo.json) "
        "yang mencakup 100% batas administrasi pemekaran wilayah baru."
    )

    add_chart_image("hasil_analisis/grafik/5_peta_geospasial_kemudahan_nego_2026.png", width_inches=6.4, caption="Grafik 5.1: Peta Tematik Geospasial Kemudahan & Kesulitan Negosiasi Wilayah PGI (Tahun 2026)")

    add_heading_2("5.1 Sebaran Spasial & Zonasi Geografis Kemudahan Negosiasi")
    add_body(
        "Peta tematik di atas (Grafik 5.1) memperlihatkan visualisasi zonasi spasial kemudahan negosiasi pada koridor operasional utama PGI (Pulau Jawa dan Bali):\n"
        "• Zona Hijau Zamrud (Sangat Mudah, Skor ≥ 70): Terkonsentrasi di wilayah penyangga Jawa Tengah bagian selatan (Cilacap, Banjarnegara), Jawa Barat Priangan (Tasikmalaya, Garut, Sukabumi), Banten (Serang), dan koridor metropolitan prime (Jakarta Pusat).\n"
        "• Zona Biru Langit (Mudah, Skor 60–69,9): Mendominasi kawasan aglomerasi Jabodetabek (Bekasi, Jakarta Selatan, Tangsel, Bogor), Surabaya, Subang, dan pusat komersial pariwisata Badung Bali.\n"
        "• Zona Kuning / Oranye (Moderat s/d Sulit, Skor 40–59,9): Meliputi jalur Pantura Jawa Tengah (Pekalongan, Tegal, Batang, Semarang), Solo Raya (Surakarta, Sukoharjo, Klaten), dan kawasan industri Purwasuka (Purwakarta, Karawang).\n"
        "• Zona Merah Tua (Sangat Sulit / Paling Alot, Skor < 40): Menjadi episentrum resistensi tawar-menawar terberat, berpusat di Kab. Bandung Barat, Kab. Grobogan, Kota Cirebon, Kab. Cianjur, Kab. Cirebon, dan Kab. Ciamis."
    )

    add_chart_image("hasil_analisis/grafik/6_analisis_karakteristik_wilayah_nego_2026.png", width_inches=6.4, caption="Grafik 5.2: Matriks Disparitas Skor Kemudahan & Profil Efisiensi Diskon Klaster Regional (Tahun 2026)")

    add_heading_2("5.2 Analisis Karakteristik Kota/Kabupaten Termudah (Top 10 Termudah)")
    add_body(
        "Sepuluh wilayah dengan Skor Kemudahan Negosiasi tertinggi pada kohort 2026 (minimal 2 transaksi) mencatatkan kombinasi luar biasa antara "
        "diskon tinggi (rata-rata 24,3%) dan durasi penutupan transaksi yang sangat kilat (rata-rata 6,7 hari kalender)."
    )

    t7_tbl = doc.add_table(rows=11, cols=6)
    format_table(t7_tbl, [1.8, 0.7, 0.9, 0.9, 1.1, 1.1])
    t7_headers = ["Wilayah (Top 10 Termudah)", "Jumlah Pengajuan", "Avg Durasi", "Avg Diskon", "Total Saving", "Skor Kemudahan"]
    for i, h in enumerate(t7_headers):
        t7_tbl.rows[0].cells[i].paragraphs[0].text = h

    t7_data = [
        ("Kab. Cilacap", "3 unit", "8,7 hari", "40,0%", "Rp 48,5 Jt", "92,2 (Sangat Mudah)"),
        ("Kota Tasikmalaya", "2 unit", "4,0 hari", "22,5%", "Rp 30,0 Jt", "76,7 (Mudah)"),
        ("Kab. Serang", "3 unit", "9,0 hari", "27,7%", "Rp 20,0 Jt", "75,2 (Mudah)"),
        ("Kab. Banjarnegara", "2 unit", "9,0 hari", "27,5%", "Rp 8,0 Jt", "75,0 (Mudah)"),
        ("Kota Jakarta Pusat", "3 unit", "7,7 hari", "24,3%", "Rp 67,0 Jt", "73,0 (Mudah)"),
        ("Kab. Sukabumi", "3 unit", "8,0 hari", "23,3%", "Rp 32,0 Jt", "71,1 (Mudah)"),
        ("Kab. Garut", "3 unit", "5,3 hari", "19,3%", "Rp 18,0 Jt", "70,2 (Mudah)"),
        ("Kab. Subang", "3 unit", "7,3 hari", "21,0%", "Rp 23,0 Jt", "69,1 (Mudah)"),
        ("Kab. Badung (Bali)", "3 unit", "6,7 hari", "19,7%", "Rp 50,0 Jt", "68,4 (Mudah)"),
        ("Kab. Bekasi", "6 unit", "5,2 hari", "17,3%", "Rp 43,0 Jt", "67,8 (Mudah)")
    ]
    for idx, r_data in enumerate(t7_data):
        row = t7_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_body(
        "Dekonstruksi Profil Karakteristik Wilayah Termudah:\n"
        "1. Fenomena 'Buyer's Market' & Kelebihan Pasokan (Oversupply) Ruko di Sub-Urban / Tier 2–3:\n"
        "   Wilayah seperti Kab. Cilacap (skor 92,2; diskon fantastis 40,0%), Kab. Banjarnegara (27,5%), Kab. Serang (27,7%), dan Kab. Subang (21,0%) dicirikan oleh maraknya pembangunan ruko baru yang tidak diimbangi lonjakan penyerapan usaha lokal. Pemilik ruko perorangan di wilayah ini menghadapi risiko properti kosong (opportunity loss). Ketika tim negosiator PGI mengajukan penawaran institusional dengan komitmen sewa jangka panjang (3–5 tahun), pemilik ruko sangat kooperatif dan rela memangkas harga puluhan persen demi mengamankan pendapatan pasif yang pasti.\n"
        "2. Ruang Tawar Lebar (High-Margin Buffer) pada Koridor Metropolitan Prime:\n"
        "   Wilayah Kota Jakarta Pusat (skor 73,0; diskon 24,3%; saving Rp 67 Jt), Kab. Badung Bali (skor 68,4; diskon 19,7%; saving Rp 50 Jt), Kota Surabaya (skor 65,3; diskon 21,5%), dan Kota Jakarta Selatan (skor 67,6; diskon 17,9%; saving Rp 76,5 Jt) memiliki harga penawaran awal ruko yang tinggi (rata-rata asking Rp 60M – Rp 87M per ruko). Tingginya patokan awal ini memberikan bantalan ruang tawar (markup buffer) yang sangat besar. Keahlian negosiator senior (terutama Bonita) mampu mengeksploitasi ruang tawar ini dengan skema multi-tahun corporate guarantor, menghasilkan nominal saving masif dalam tempo di bawah 8 hari.\n"
        "3. Kematangan Literasi Kontrak Komersial & Transaksi Tanpa Birokrasi Keluarga:\n"
        "   Pemilik ruko di wilayah urban dan sentra niaga penyangga sudah terbiasa dengan transaksi legalitas korporasi, pemotongan pajak PPh/PPN, dan transfer perbankan termin. Pengambilan keputusan sewa bersifat langsung (single decision-maker), sehingga closing tuntas tanpa jeda birokrasi keluarga."
    )

    add_heading_2("5.3 Analisis Karakteristik Kota/Kabupaten Tersulit & Paling Alot (Top 10 Tersulit)")
    add_body(
        "Sebaliknya, sepuluh wilayah dengan skor terendah memperlihatkan resistensi negosiasi yang sangat tinggi. "
        "Rata-rata diskon yang berhasil diperoleh hanya 11,2% dengan proses tawar-menawar yang berlarut-larut mencapai rata-rata 20,5 hari kalender "
        "(hampir dua kali lipat rata-rata nasional 11,5 hari)."
    )

    t8_tbl = doc.add_table(rows=11, cols=6)
    format_table(t8_tbl, [1.8, 0.7, 0.9, 0.9, 1.1, 1.1])
    t8_headers = ["Wilayah (Top 10 Tersulit)", "Jumlah Pengajuan", "Avg Durasi", "Avg Diskon", "Total Saving", "Skor Kemudahan"]
    for i, h in enumerate(t8_headers):
        t8_tbl.rows[0].cells[i].paragraphs[0].text = h

    t8_data = [
        ("Kab. Bandung Barat", "3 unit", "24,3 hari", "4,7%", "Rp 5,2 Jt", "19,0 (Paling Alot)"),
        ("Kab. Grobogan", "2 unit", "34,0 hari", "17,5%", "Rp 13,5 Jt", "20,0 (Sangat Sulit)"),
        ("Kota Cirebon", "2 unit", "17,0 hari", "2,5%", "Rp 1,7 Jt", "28,3 (Sulit)"),
        ("Kab. Cianjur", "8 unit", "22,0 hari", "8,9%", "Rp 26,0 Jt", "28,5 (Sulit)"),
        ("Kab. Cirebon", "4 unit", "18,2 hari", "9,2%", "Rp 12,0 Jt", "35,2 (Sulit)"),
        ("Kab. Ciamis", "3 unit", "17,3 hari", "9,3%", "Rp 7,0 Jt", "36,9 (Sulit)"),
        ("Kab. Blora", "3 unit", "20,3 hari", "16,0%", "Rp 30,0 Jt", "40,8 (Alot)"),
        ("Kab. Banyumas", "4 unit", "20,0 hari", "17,2%", "Rp 27,0 Jt", "43,0 (Alot)"),
        ("Kab. Purwakarta", "6 unit", "17,8 hari", "15,2%", "Rp 37,0 Jt", "43,8 (Alot)"),
        ("Kota Surakarta", "3 unit", "14,7 hari", "12,3%", "Rp 23,5 Jt", "45,3 (Alot)")
    ]
    for idx, r_data in enumerate(t8_data):
        row = t8_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_body(
        "Dekonstruksi Profil Karakteristik Wilayah Paling Alot:\n"
        "1. Cengkeraman Pusat Niaga Tradisional & Koridor Heritage Matang (Mature Commercial Strips):\n"
        "   Kota Cirebon (skor 28,3; diskon miris hanya 2,5%; durasi 17,0 hari), Kab. Cirebon (skor 35,2; diskon 9,2%), Kota Surakarta/Solo (skor 45,3; durasi 14,7 hari; asking Rp 85M), Kab. Klaten (skor 45,7; diskon 8,0%), dan Kota Tangerang (skor 48,5; diskon 11,8% di 11 cabang) merupakan kawasan niaga lama dengan lalu lintas konsumen yang padat dan mapan. Pemilik ruko umumnya adalah keluarga saudagar atau investor tradisional berkocek tebal yang memiliki 'holding power' keuangan sangat kuat. Mereka tidak memerlukan likuiditas mendesak dan yakin propertinya akan tetap laku, sehingga menerapkan prinsip harga mati (rigid price mindset) dan menolak negosiasi formal korporasi.\n"
        "2. Hambatan Kultural Aset Warisan & Kepemilikan Majemuk di Wilayah Rural/Agrikultur:\n"
        "   Kab. Bandung Barat (skor 19,0; durasi 24,3 hari; diskon hanya 4,7% — dinobatkan sebagai Wilayah Paling Alot se-Indonesia), Kab. Grobogan (skor 20,0; durasi rekor terlama 34,0 hari), Kab. Cianjur (skor 28,5; durasi 22,0 hari di 8 cabang), Kab. Ciamis (skor 36,9; durasi 17,3 hari), dan Kab. Blora (skor 40,8; durasi 20,3 hari) memiliki problem kepemilikan aset keluarga/warisan (family-owned inheritance assets). Proses penurunan harga memerlukan konsensus seluruh anggota keluarga besar yang sering kali tidak berada di satu kota. Rantai persetujuan yang berbelit ini menyedot waktu negosiasi hingga berminggu-minggu tanpa menghasilkan diskon berarti.\n"
        "3. Risiko Finansial 'Cost of Delay':\n"
        "   Alotnya proses di wilayah-wilayah ini mengunci modal ekspansi PGI. Keterlambatan tawar-menawar selama 20–34 hari secara otomatis menunda fase renovasi fisik kontraktor dan menunda perolehan omzet cabang baru."
    )

    # SUB-BAB BARU: UJI EMPIRIS KOTA VS KABUPATEN
    add_heading_2("5.4 Uji Empiris: Benarkah Wilayah Kota Lebih Mudah vs Kabupaten/Desa Lebih Alot?")
    add_body(
        "Untuk memverifikasi hipotesis manajemen mengenai adanya kesenjangan (*dichotomy*) kemudahan negosiasi antara wilayah perkotaan (Kota/Urban) "
        "dan wilayah kabupaten/daerah/pedesaan (Kabupaten/Rural), dilakukan komparasi empiris terhadap 275 cabang aktif tahun 2026:"
    )

    t_kotakab = doc.add_table(rows=6, cols=4)
    format_table(t_kotakab, [1.8, 1.4, 1.5, 1.8])
    t_kk_headers = ["Dimensi Metrik Negosiasi", "Kota (Urban)", "Kabupaten (Daerah/Rural)", "Kesenjangan & Temuan Data"]
    for i, h in enumerate(t_kk_headers):
        t_kotakab.rows[0].cells[i].paragraphs[0].text = h

    t_kk_data = [
        ("Cakupan & Jumlah Pengajuan", "26 Kota | 108 Pengajuan (39,3%)", "50 Kabupaten | 167 Pengajuan (60,7%)", "Mayoritas ekspansi PGI berada di Kabupaten (60,7%)."),
        ("Rerata Harga Penawaran Awal", "Rp 55,28 Juta / tahun", "Rp 36,66 Juta / tahun", "Tarif sewa ruko Kota 50,8% lebih mahal dibanding Kabupaten."),
        ("Rerata Durasi Negosiasi", "10,3 Hari (Median 8,0 hr)", "12,2 Hari (Median 8,0 hr)", "Kota 1,9 hari lebih cepat closing dibanding Kabupaten."),
        ("Rerata Efisiensi Diskon Sewa", "15,40% (Median 15,0%)", "15,69% (Median 16,0%)", "Persentase diskon relatif seimbang (~15,5%)."),
        ("Konsentrasi Kasus Alot (> 14 hr)", "17,6% transaksi alot (19 cabang)", "20,4% transaksi alot (34 cabang)", "80% dari Top 10 wilayah paling alot adalah KABUPATEN.")
    ]
    for idx, r_data in enumerate(t_kk_data):
        row = t_kotakab.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_body(
        "Kesimpulan Uji Empiris Kota vs Kabupaten:\n"
        "1. Hipotesis Terbukti Sangat Kuat pada Dimensi Resistensi & Kelambatan (Alot):\n"
        "   Fakta membuktikan bahwa 8 dari 10 wilayah paling alot dan berdurasi terlama di Indonesia berasal dari KABUPATEN (Bandung Barat, Grobogan, Cianjur, Kab. Cirebon, Ciamis, Blora, Banyumas, dan Purwakarta). Total 34 cabang di kabupaten-kabupaten ini tersendat negosiasi alot (> 14 hari) dengan rata-rata realisasi mencapai 20,5 hari. Hal ini menegaskan bahwa karakter kepemilikan ruko di daerah/kabupaten (warisan keluarga majemuk, musyawarah berbelit, dan holding power tanah pasif) merupakan sumber resistensi negosiasi terbesar bagi PGI.\n"
        "2. Anomali Penting 1: Tidak Seluruh Kabupaten Alot (Faktor Oversupply Ruko Sub-Urban):\n"
        "   Data membuktikan terdapat kabupaten yang justru menduduki peringkat teratas termudah, seperti Kab. Cilacap (diskon 40,0%, 3 pengajuan), Kab. Serang (diskon 27,7%, 3 pengajuan), Kab. Banjarnegara (diskon 27,5%, 2 pengajuan), Kab. Garut (durasi 5,3 hari, 3 pengajuan), Kab. Subang (diskon 21,0%, 3 pengajuan), dan Kab. Bekasi (durasi 5,2 hari, 6 pengajuan). Di wilayah-wilayah ini terjadi fenomena kelebihan pasokan ruko (oversupply), sehingga pemilik ruko perorangan sangat fleksibel demi mengamankan kepastian sewa jangka panjang korporasi PGI.\n"
        "3. Anomali Penting 2: Kota-Kota Niaga Tradisional yang Sangat Alot:\n"
        "   Wilayah Kota tidak secara otomatis mudah. Kota Cirebon (skor 28,3; diskon miris hanya 2,5%; 2 pengajuan) dan Kota Surakarta/Solo (skor 45,3; durasi 14,7 hari; asking Rp 85 M) masuk jajaran paling alot karena berstatus pusat niaga tradisional heritage (Pecinan & sentra batik keraton) di mana pemilik properti adalah saudagar tua bermodal kuat yang pantang menurunkan harga."
    )

    add_heading_2("5.5 Profil Komparasi 11 Klaster Regional Negosiasi")
    add_body(
        "Pengelompokan 275 cabang tahun 2026 ke dalam 11 klaster geografis memperlihatkan kontribusi strategis dan peta risiko regional yang sangat kontras:"
    )

    t_klaster = doc.add_table(rows=12, cols=7)
    format_table(t_klaster, [1.7, 0.7, 1.0, 1.0, 0.7, 0.7, 0.7])
    t_k_headers = ["Klaster Geografis", "Jumlah Pengajuan", "Total Sewa Deal", "Total Saving", "Diskon %", "Durasi", "Lead Tim"]
    for i, h in enumerate(t_k_headers):
        t_klaster.rows[0].cells[i].paragraphs[0].text = h

    t_k_data = [
        ("Jawa Tengah Selatan", "9 unit", "Rp 238,5 Jt", "Rp 78,0 Jt", "22,56%", "13,1 hr", "Mirza (5)"),
        ("Banten Barat", "13 unit", "Rp 392,0 Jt", "Rp 96,0 Jt", "20,08%", "11,5 hr", "Mirza (7)"),
        ("Jabodetabek Puncak", "82 unit", "Rp 3.341,6 Jt", "Rp 719,8 Jt", "16,09%", "9,9 hr", "Bonita (47)"),
        ("Jawa Barat Priangan & Sel.", "27 unit", "Rp 793,0 Jt", "Rp 160,0 Jt", "15,48%", "14,0 hr", "Mirza (14)"),
        ("Jawa Timur", "9 unit", "Rp 380,5 Jt", "Rp 74,5 Jt", "15,22%", "8,3 hr", "Bonita (7)"),
        ("Jawa Tengah Pantura", "36 unit", "Rp 1.094,0 Jt", "Rp 220,3 Jt", "14,75%", "11,6 hr", "Mirza (19)"),
        ("Bandung Raya", "23 unit", "Rp 1.025,0 Jt", "Rp 203,2 Jt", "14,65%", "11,8 hr", "Bonita (12)"),
        ("Solo Raya & DIY", "16 unit", "Rp 595,5 Jt", "Rp 95,5 Jt", "14,44%", "12,2 hr", "Mirza (11)"),
        ("Bali", "12 unit", "Rp 646,0 Jt", "Rp 124,0 Jt", "14,33%", "9,0 hr", "Bonita (7)"),
        ("Jabar Pantura & Purwasuka", "30 unit", "Rp 885,8 Jt", "Rp 141,2 Jt", "13,63%", "15,3 hr", "Bonita (15)"),
        ("Wilayah Luar Jawa/Lainnya", "18 unit", "Rp 664,0 Jt", "Rp 123,0 Jt", "14,72%", "9,8 hr", "Mirza (10)")
    ]
    for idx, r_data in enumerate(t_k_data):
        row = t_klaster.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_body(
        "Insight Klaster Regional Utama:\n"
        "• Jabodetabek Puncak sebagai Pusat Kontribusi: Menyerap 82 cabang (29,8% nasional) dan menyumbang penghematan terbesar Rp 719,8 Juta (35,4% total saving 2026) dengan durasi sangat efisien (9,9 hari).\n"
        "• Jawa Tengah Selatan & Banten Barat sebagai Juara Efisiensi Diskon: Kedua klaster ini mencatat diskon rata-rata tertinggi nasional (> 20%), membuktikan tingginya elastisitas harga ruko di koridor tersebut.\n"
        "• Jawa Barat Pantura & Purwasuka sebagai Klaster Paling Berisiko: Membukukan durasi terlama nasional (15,3 hari) dan efisiensi diskon terendah (13,63%), dipicu oleh tingginya kasus ruko alot di Cirebon dan Purwakarta.\n"
        "• Bali sebagai Pasar Bernilai Tinggi: Memiliki rata-rata sewa ruko tertinggi (Rp 53,8 Juta per unit deal), namun mampu dituntaskan dalam waktu 9,0 hari dengan total saving Rp 124 Juta."
    )

    add_heading_2("5.6 Rekomendasi Taktikal & SOP Penugasan Negosiator di Wilayah Alot")
    add_callout(
        "PEDOMAN OPERASIONAL MITIGASI WILAYAH ALOT:\n"
        "1. Kebijakan Cut-off 14 Hari & Lokasi Cadangan (Dual-Sourcing):\n"
        "   Manajemen wajib memberlakukan SOP batas maksimal proses negosiasi 14 hari kalender di wilayah alot (Bandung Barat, Grobogan, Cirebon, Cianjur). Jika pada hari ke-14 diskon yang ditawarkan pemilik masih < 5%, surveyor diwajibkan menyodorkan Ruko B dan C sebagai posisi tawar pengganti (BATNA - Best Alternative to a Negotiated Agreement).\n"
        "2. Skema Penugasan Spesialis (Persona-Matching):\n"
        "   • Bonita dialokasikan khusus menangani koridor bernilai sewa tinggi (> Rp 50 Juta) dan wilayah perkotaan besar (Jabodetabek, Bali, Surabaya, Bandung Kota) untuk memaksimalkan nominal saving jutaan rupiah.\n"
        "   • Mirza difokuskan pada wilayah satelit, Pantura, dan ekspansi cepat Jawa Tengah/Jawa Barat yang menuntut kecepatan closing tinggi (< 10 hari) guna mencegah kebuntuan negosiasi.\n"
        "3. Pemanfaatan Skema Corporate Guarantor Multi-Tahun:\n"
        "   Untuk melunakkan tuan tanah tradisional ber-holding power kuat di kota-kota niaga lama, gunakan senjata legalitas: komitmen sewa pasti 3–5 tahun, klausul perbaikan renovasi ruko yang dibiayai PGI (meningkatkan nilai aset ruko mereka), dan pembayaran sewa tahunan di muka bergaransi institusi korporasi resmi.",
        title="SOP Taktikal Negosiasi Wilayah PGI",
        bg_hex="F0FDF4", border_hex="16A34A"
    )

    # ==========================================
    # BAB 6: KONTRAKTOR & SLA RENOVASI
    # ==========================================
    add_heading_1("BAB 6: EVALUASI KONTRAKTOR & MITIGASI RISIKO SLA RENOVASI")
    add_body(
        "Total komitmen anggaran renovasi fisik pada 1.028 cabang ruko PGI mencapai Rp 126,05 Miliar. "
        "Meskipun target Service Level Agreement (SLA) renovasi secara formal ditetapkan rata-rata 21,5 hari kerja, "
        "tingkat kepatuhan kontraktor secara keseluruhan mengalami penurunan konsisten yang sangat mengkhawatirkan: "
        "dari 68,1% (2024), turun ke 64,2% (2025), dan anjlok menjadi 53,8% (2026). "
        "Penurunan kepatuhan sebesar 14,3 poin persentase ini memperpanjang masa tunggu pembukaan cabang baru dan menunda realisasi perputaran omzet kas perusahaan."
    )

    add_heading_2("6.1 Peta Risiko Vendor Komparatif (Side-by-Side Vendor Risk Quadrant: 2025 vs 2026)")
    add_body(
        "Sesuai arahan audit operasional (DOOM UPC Poin A), evaluasi kontraktor difokuskan secara murni pada Peta Risiko Vendor "
        "(Vendor Risk Quadrant) tanpa distraksi kartu pengujian di luar kuadran. "
        "Untuk melacak evolusi kinerja dan pergeseran posisi setiap kontraktor secara objektif, disusun analisis komparatif "
        "antara Tahun 2025 (436 cabang) dan Tahun 2026 (275 cabang) berdasarkan dua sumbu utama:\n"
        "1. Volume Pengerjaan (Sumbu X): Jumlah cabang yang dikerjakan per tahun (ambang batas: 50 unit di 2025, 40 unit di 2026).\n"
        "2. Tingkat Kepatuhan SLA (Sumbu Y): Persentase proyek yang diserahterimakan tepat waktu (ambang batas sehat: 70,0%)."
    )

    add_chart_image("hasil_analisis/grafik/8_peta_risiko_vendor_komparasi_2025_2026.png", width_inches=6.4,
                    caption="Grafik 6.1: Peta Risiko Vendor Berdampingan: Perubahan Posisi dan Performa Kontraktor (Tahun 2025 vs 2026)")

    add_body(
        "Matriks berikut menyajikan komparasi kuantitatif performa, durasi pengerjaan fisik, tingkat kepatuhan SLA, "
        "dan pergeseran kuadran untuk seluruh mitra kontraktor antara tahun 2025 dan 2026:"
    )

    t_comp_vendor = doc.add_table(rows=8, cols=8)
    format_table(t_comp_vendor, [1.6, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7, 1.2])
    t_cv_headers = ["Mitra Kontraktor", "Unit '25", "SLA '25", "Unit '26", "SLA '26", "Delta SLA", "Dur '26", "Migrasi Kuadran"]
    for i, h in enumerate(t_cv_headers):
        t_comp_vendor.rows[0].cells[i].paragraphs[0].text = h

    t_cv_data = [
        ("Sarjono", "109", "88,1%", "73", "89,0%", "+1,0 pt", "14,9 hr", "Kuadran I -> Kuadran I (Stable Champion)"),
        ("CV. Rizki Mitra Abadi", "74", "44,6%", "54", "66,7%", "+22,1 pt", "21,7 hr", "Kuadran II -> Mendekati Kuadran I"),
        ("Edwin", "66", "69,7%", "52", "61,5%", "-8,2 pt", "21,4 hr", "Kuadran II -> Kuadran II (Moderat)"),
        ("Trimo*", "74", "41,9%", "53", "15,1%", "-26,8 pt", "34,5 hr", "Kuadran II -> Kuadran II (Beban Bali)"),
        ("Teguh Karyanto", "2", "0,0%", "30", "20,0%", "+20,0 pt", "31,3 hr", "Kuadran III -> Kuadran III (Alot)"),
        ("Sendy", "70", "98,6%", "3", "66,7%", "-31,9 pt", "18,3 hr", "Kuadran I -> Kuadran III (Volume Turun)"),
        ("CV Cahaya Kemakmuran", "39", "15,4%", "0", "-", "-", "-", "Kuadran III -> Dinonaktifkan (Off)")
    ]
    for idx, r_data in enumerate(t_cv_data):
        row = t_comp_vendor.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx in [0, 1]:
            row.cells[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(5, 150, 105)
        elif idx == 3:
            row.cells[5].paragraphs[0].runs[0].font.color.rgb = RGBColor(220, 38, 38)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_body("Catatan: (*) Penurunan SLA Trimo di 2026 sangat dipengaruhi oleh penugasan khusus ekspansi 12 cabang di wilayah Bali.", italic=True)

    add_heading_2("6.2 Peta Vektor Lintasan & Migrasi Risiko Kontraktor (2025 ke 2026)")
    add_body(
        "Untuk memahami arah dinamika setiap mitra, Grafik 6.2 memetakan vektor pergerakan dari posisi tahun 2025 "
        "(lingkaran putih putus-putus) menuju posisi realisasi tahun 2026 (titik tebal berpanah):"
    )

    add_chart_image("hasil_analisis/grafik/9_vektor_migrasi_risiko_vendor_2025_2026.png", width_inches=6.4,
                    caption="Grafik 6.2: Peta Vektor Lintasan Migrasi Risiko Vendor: Melacak Dinamika Kapasitas Eksekusi (2025 -> 2026)")

    add_body(
        "Dekomposisi Dinamika Pergerakan Kontraktor:\n"
        "1. Sarjono (The Stable Core Champion):\n"
        "   Sarjono adalah mitra paling konsisten dan tangguh di jaringan PGI. Meskipun volume sedikit disesuaikan dari 109 unit (2025) "
        "   menjadi 73 unit (2026), kepatuhan SLA-nya justru meningkat menjadi 89,0% dengan rata-rata durasi penyelesaian kilat 14,9 hari. "
        "   Sarjono kokoh berdiri di Kuadran I sebagai tulang punggung utama ekspansi fisik PGI.\n"
        "2. CV Rizki Mitra Abadi (The Turnaround Improver):\n"
        "   Mencatatkan lonjakan performa paling impresif se-Indonesia. Kepatuhan SLA melejit drastis sebesar +22,1 poin persentase "
        "   (dari 44,6% di 2025 menjadi 66,7% di 2026 pada 54 cabang) dengan rata-rata durasi 21,7 hari. Vektor pergerakannya mengarah "
        "   tajam ke atas menuju gerbang Kuadran I (Core Champion).\n"
        "3. Edwin (Moderate & Stabil):\n"
        "   Menangani 52 cabang di 2026 dengan tingkat kepatuhan 61,5% dan durasi rata-rata 21,4 hari. Edwin mempertahankan kapasitasnya "
        "   sebagai kontraktor lapis kedua yang stabil, meskipun sedikit di bawah garis ambang batas 70%.\n"
        "4. Teguh Karyanto & Dimas Andri Sulistyo (Ekspansi Berisiko Tinggi):\n"
        "   Teguh Karyanto mengalami lonjakan volume drastis dari hanya 2 cabang (2025) menjadi 30 cabang (2026). Namun, lonjakan beban ini "
        "   tidak diimbangi kesiapan kapasitas, menghasilkan durasi molor rata-rata 31,3 hari dan kepatuhan rendah 20,0%. "
        "   Sementara itu, Dimas Andri Sulistyo sebagai mitra baru menangani 9 cabang dengan durasi 30,4 hari dan SLA hanya 11,1%. "
        "   Keduanya terkonsentrasi di Kuadran III (Underperformers).\n"
        "5. Sendy & CV Cahaya:\n"
        "   Sendy yang pada 2025 merupakan Core Champion (70 unit, SLA 98,6%) mengalami penciutan kuota drastis menjadi hanya 3 unit di 2026. "
        "   Sedangkan CV Cahaya Kemakmuran Sejahtera dinonaktifkan total (0 unit) setelah performa buruk di 2025 (SLA 15,4%)."
    )

    add_heading_2("6.3 Analisis Realitas Lapangan: Beban Penugasan Khusus Wilayah Bali pada Kontraktor Trimo")
    add_body(
        "Berdasarkan visualisasi kuadran, kontraktor Trimo tampak terpuruk di Kuadran II dengan SLA anjlok menjadi 15,1% dan durasi rata-rata membengkak ke 34,5 hari. "
        "Namun, audit operasional mendalam mengungkap faktor fundamental yang melandasi anomali ini: "
        "Trimo adalah satu-satunya kontraktor yang bersedia mengambil risiko dan menanggung penugasan ekspansi cabang di Provinsi Bali (12 cabang di 2026), "
        "sedangkan seluruh kontraktor lainnya (Sarjono, Sendy, Edwin, Teguh) menolak atau belum bersedia mengambil proyek di wilayah Bali karena tantangan jarak dan supervisi lintas pulau."
    )

    t_bali_tbl = doc.add_table(rows=4, cols=5)
    format_table(t_bali_tbl, [2.0, 1.0, 1.2, 1.2, 1.4])
    t_b_headers = ["Segmentasi Wilayah Trimo 2026", "Jumlah Unit", "Rerata Durasi", "Kepatuhan SLA", "Dampak Operasional"]
    for i, h in enumerate(t_b_headers):
        t_bali_tbl.rows[0].cells[i].paragraphs[0].text = h

    t_b_data = [
        ("Proyek Khusus Wilayah Bali", "12 cabang (22,6%)", "43,9 hari (Med 41,5 hr)", "0,0% (0 tercapai)", "Logistik fery, pasokan antarpulau, izin Banjar"),
        ("Proyek Regular Luar Bali (Jawa)", "41 cabang (77,4%)", "31,8 hari (Med 32,0 hr)", "19,5% (8 tercapai)", "Problem internal kapasitas mandor & sub-kon"),
        ("TOTAL PORTOFOLIO TRIMO 2026", "53 cabang (100%)", "34,5 hari (Med 35,0 hr)", "15,1% (8 tercapai)", "SLA agregat terdistorsi tajam oleh 12 proyek Bali")
    ]
    for idx, r_data in enumerate(t_b_data):
        row = t_bali_tbl.rows[idx+1]
        for c_idx, val in enumerate(r_data):
            row.cells[c_idx].paragraphs[0].text = val
        if idx == 2:
            for c_idx in range(len(r_data)):
                row.cells[c_idx].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_body(
        "Faktor-Faktor Lapangan Penghambat Eksekusi di Bali:\n"
        "1. Hambatan Logistik & Distribusi Antarpulau:\n"
        "   Seluruh material branding khusus, partisi besi, kaca securit, dan brankas gadai harus dikirim dari gudang Jawa Timur/Surabaya via penyeberangan fery Ketapang–Gilimanuk. "
        "   Antrean pelabuhan, cuaca selat, dan moda transportasi darat lintas selat menambah lead time pengiriman material 7–14 hari kerja dibanding proyek Pulau Jawa.\n"
        "2. Regulasi Adat Banjar & Dinamika Upacara Keagamaan Lokal:\n"
        "   Aktivitas renovasi fisik di Denpasar, Badung, Tabanan, Gianyar, dan Buleleng terikat pada aturan banjar adat setempat. "
        "   Pengerjaan sering kali harus dihentikan sementara (silent days) saat berlangsungnya upacara keagamaan (Rahinan, Piodalan, Ngaben massal), "
        "   sehingga hari kerja efektif jauh lebih pendek dari hari kalender.\n"
        "3. Soliter Tanpa Kompetitor (Single Contractor Risk):\n"
        "   Karena Sarjono, Sendy, dan Edwin belum bersedia masuk ke pasar Bali, PGI terpaksa membebankan seluruh 12 titik gerai baru Bali kepada Trimo secara borongan. "
        "   Hal ini mengakibatkan bottleneck sumber daya mandor lokal yang sangat parah di Bali.\n"
        "4. Kinerja Trimo di Luar Bali:\n"
        "   Apabila 12 cabang Bali dieksklusi, durasi rata-rata Trimo di Pulau Jawa adalah 31,8 hari. Angka ini memang masih memerlukan peningkatan disiplin SLA, "
        "   namun membuktikan bahwa keterpurukan Trimo ke angka 34,5 hari dan 15,1% SLA sangat didistorsi oleh kerelaannya mengambil proyek perintis di Pulau Dewata."
    )

    add_heading_2("6.4 Rekomendasi Manajerial & Solusi Penanganan Wilayah Khusus (Bali)")
    add_callout(
        "SOLUSI MANAJERIAL PENYEHATAN KONTRAKTOR & STRATEGI WILAYAH BALI:\n"
        "1. Penyesuaian SLA Regional Khusus Bali (35–40 Hari Kalender):\n"
        "   Target SLA renovasi nasional (21 hari) tidak realistis diterapkan untuk wilayah kepulauan seperti Bali. "
        "   Manajemen wajib menerbitkan SK penyesuaian SLA resmi khusus regional Bali menjadi 35–40 hari kerja guna mengakomodasi faktor fery dan adat.\n"
        "2. Skema Insentif Mobilisasi & Berbagi Beban (Burden Sharing):\n"
        "   Manajemen tidak boleh membiarkan Trimo menanggung beban Bali seorang diri. Berikan insentif biaya mobilisasi atau penyesuaian tarif "
        "   kepada Sarjono dan CV Rizki Mitra Abadi agar bersedia membuka cabang operasional di Bali, sehingga tercipta kompetisi sehat.\n"
        "3. Pembentukan Hub Material Lokal di Denpasar:\n"
        "   Kerjasamakan pengadaan bahan baku standar (cat, semen, keramik, baja ringan) langsung dengan distributor lokal di Denpasar "
        "   untuk memotong ketergantungan pengiriman dari Jawa Timur.\n"
        "4. Optimalisasi Redistribusi Proyek Pulau Jawa:\n"
        "   Untuk proyek renovasi baru di Pulau Jawa, pangkas kuota Trimo dan alihkan kepada Sarjono serta CV Rizki Mitra Abadi yang terbukti memiliki kapasitas eksekusi tinggi.",
        title="Rekomendasi Strategis Penyehatan Vendor & Mitigasi Bali",
        bg_hex="EFF6FF", border_hex="2563EB"
    )

    # ==========================================
    # BAB 7: KESIMPULAN & REKOMENDASI
    # ==========================================
    add_heading_1("BAB 7: KESIMPULAN & REKOMENDASI MANAJERIAL STRATEGIS")
    add_body(
        "Berdasarkan hasil analisis komprehensif pada dataset bersih 1.028 cabang, Divisi Bisnis mengajukan 4 pilar rekomendasi strategis:"
    )

    add_heading_2("1. Redistribusi Kuota & Manajemen Penalti Kontraktor")
    add_body(
        "• Alihkan minimal 30%–50% alokasi proyek baru dari Trimo, CV CKS, dan Teguh Karyanto kepada Sarjono dan Sendy.\n"
        "• Terapkan klausul denda keterlambatan (liquidated damages) sebesar 1% per hari keterlambatan terhadap SLA yang disepakati guna menekan pembengkakan waktu renovasi di 2026."
    )

    add_heading_2("2. Standarisasi SOP Negosiasi & Batas Waktu Cut-Off 14 Hari")
    add_body(
        "• Tetapkan batas waktu (cut-off) negosiasi maksimal 14 hari kalender. Data membuktikan bahwa negosiasi yang berlarut-larut "
        "(korelasi r = -0,104) tidak menghasilkan diskon yang lebih besar dan hanya membuang waktu sewa.\n"
        "• Standardisasi formula dan teknik tawar Bonita (diskon rata-rata 18,36% pada ruko premium) sebagai modul pelatihan internal bagi negosiator junior."
    )

    add_heading_2("3. Penerapan Peta Geospasial pada Tahap Feasibility Study")
    add_body(
        "• Gunakan indeks kemudahan negosiasi wilayah (Tabel 5.1 & 5.2) dalam penentuan target sewa lokasi baru. "
        "Untuk wilayah alot seperti Kab. Bandung Barat, Kab. Grobogan, dan Kota Cirebon, manajemen perlu menyiapkan rencana kontingensi "
        "atau mencari alternatif ruko lapis kedua lebih awal."
    )

    add_heading_2("4. Integrasi Validasi Form Database & Penghapusan Redudansi")
    add_body(
        "• Kunci input form sistem internal agar kolom harga sewa deal tidak meng-overwrite harga penawaran awal pemilik ruko, "
        "sehingga penghematan riil tercatat otomatis tanpa memerlukan rekonsiliasi lembar kerja terpisah.\n"
        "• Hapus kolom-kolom statis tidak bernilai (status_event, renovation_type, payment_stage) guna menyederhanakan arsitektur data ekspansi."
    )

    # ==========================================
    # LEMBAR PENGESAHAN
    # ==========================================
    doc.add_page_break()
    add_heading_1("LEMBAR PENGESAHAN DOKUMEN")
    doc.add_paragraph().paragraph_format.space_before = Pt(20)
    
    sign_table = doc.add_table(rows=4, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sign_table, color="CBD5E1")
    
    sign_headers = ["Disusun Oleh,", "Disetujui & Diketahui Oleh,"]
    for i, h in enumerate(sign_headers):
        sign_table.rows[0].cells[i].paragraphs[0].text = h
        sign_table.rows[0].cells[i].paragraphs[0].runs[0].font.bold = True
        set_cell_background(sign_table.rows[0].cells[i], "F1F5F9")
        set_cell_margins(sign_table.rows[0].cells[i], top=100, bottom=100, left=140, right=140)

    for r in range(1, 4):
        for c in range(2):
            set_cell_background(sign_table.rows[r].cells[c], "FFFFFF")
            set_cell_margins(sign_table.rows[r].cells[c], top=80, bottom=80, left=140, right=140)

    sign_table.rows[1].cells[0].paragraphs[0].text = "\n\n( Tanda Tangan )\n"
    sign_table.rows[1].cells[1].paragraphs[0].text = "\n\n( Tanda Tangan )\n"
    
    p_penyusun = sign_table.rows[2].cells[0].paragraphs[0]
    p_penyusun.text = "Mukhammad Rekza Mufti"
    p_penyusun.runs[0].font.bold = True
    p_penyusun.runs[0].font.underline = True
    
    p_penyetuju = sign_table.rows[2].cells[1].paragraphs[0]
    p_penyetuju.text = "Head of Business Development"
    p_penyetuju.runs[0].font.bold = True
    p_penyetuju.runs[0].font.underline = True

    sign_table.rows[3].cells[0].paragraphs[0].text = "Data Analyst — Divisi Bisnis\nPusat Gadai Indonesia"
    sign_table.rows[3].cells[1].paragraphs[0].text = "Divisi Bisnis & Ekspansi Cabang\nPusat Gadai Indonesia"

    output_filename = "Laporan_Analisis_UPC_PGI_v3_Data_Terbaru_2026.docx"
    doc.save(output_filename)
    
    # Simpan juga ke dalam folder Dokumen/
    dir_kerja = os.path.dirname(os.path.abspath(__file__))
    dokumen_dir = os.path.join(dir_kerja, "Dokumen")
    os.makedirs(dokumen_dir, exist_ok=True)
    doc.save(os.path.join(dokumen_dir, output_filename))
    print(f"Report saved successfully as '{output_filename}' (in root and 'Dokumen/' folder)!")

if __name__ == "__main__":
    create_report()
