#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================================================
BUILD EXECUTIVE PRESENTATION (PPTX) - PUSAT GADAI INDONESIA (PGI)
====================================================================================================
Menghasilkan presentasi PowerPoint eksekutif berstandar korporat modern (16:9 widescreen)
yang merangkum seluruh hasil analisis data pembukaan cabang (UPC), evaluasi negosiator,
korelasi durasi vs diskon, peta geospasial, vendor risk quadrant, dan sistem cerdas simulasi.
====================================================================================================
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CHART_DIR = os.path.join(BASE_DIR, "hasil_analisis", "grafik")
OUTPUT_PPTX = os.path.join(BASE_DIR, "Laporan_Eksekutif_Analisis_UPC_PGI_2026.pptx")

# --- COLOR PALETTE DEFINITION ---
NAVY_PRIMARY = RGBColor(15, 23, 42)      # #0F172A Dark Slate
NAVY_SECONDARY = RGBColor(30, 41, 59)    # #1E293B Slate 800
BLUE_ACCENT = RGBColor(37, 99, 235)      # #2563EB Royal Blue
BLUE_LIGHT = RGBColor(239, 246, 255)     # #EFF6FF Ice Blue
GOLD_ACCENT = RGBColor(217, 119, 6)      # #D97706 Warm Amber
GREEN_SUCCESS = RGBColor(16, 185, 129)   # #10B981 Emerald
RED_DANGER = RGBColor(220, 38, 38)       # #DC2626 Coral Red
BG_LIGHT = RGBColor(248, 250, 252)       # #F8FAFC Off White
TEXT_DARK = RGBColor(30, 41, 59)         # #1E293B Slate 800
TEXT_MUTED = RGBColor(100, 116, 139)     # #64748B Slate 500
WHITE = RGBColor(255, 255, 255)          # White
CARD_BORDER = RGBColor(226, 232, 240)    # #E2E8F0 Light Gray

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="PT PUSAT GADAI INDONESIA  |  DIVISI BISNIS & EXPANSION"):
        # Header banner box
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = NAVY_PRIMARY
        top_bar.line.color.rgb = NAVY_PRIMARY

        # Category / Subtitle
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.7), Inches(0.3))
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = category_text.upper()
        p_c.font.size = Pt(9.5)
        p_c.font.bold = True
        p_c.font.color.rgb = GOLD_ACCENT

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.7), Inches(0.65))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(20)
        p_t.font.bold = True
        p_t.font.color.rgb = WHITE

        # Gold accent line
        gold_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.1), Inches(13.333), Inches(0.04))
        gold_line.fill.solid()
        gold_line.fill.fore_color.rgb = GOLD_ACCENT
        gold_line.line.color.rgb = GOLD_ACCENT

        # Background tint for content
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.14), Inches(13.333), Inches(6.36))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_LIGHT
        bg.line.fill.background()

    def add_card(slide, left, top, width, height, bg_color=WHITE, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)
        return card

    # =========================================================================
    # SLIDE 1: COVER SLIDE
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY_PRIMARY
    bg1.line.fill.background()

    # Gold decorative top bar
    bar_gold = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.15))
    bar_gold.fill.solid()
    bar_gold.fill.fore_color.rgb = GOLD_ACCENT
    bar_gold.line.fill.background()

    # Sub-tag
    tag_box = slide1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.3), Inches(0.5))
    tf_tag = tag_box.text_frame
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = "PT PUSAT GADAI INDONESIA  •  DIVISI BISNIS & EXPANSION  •  LAPORAN RESMI 2026"
    p_tag.font.size = Pt(12)
    p_tag.font.bold = True
    p_tag.font.color.rgb = GOLD_ACCENT

    # Main Title
    t_box = slide1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(2.2))
    tf_main = t_box.text_frame
    tf_main.word_wrap = True
    p_main = tf_main.paragraphs[0]
    p_main.text = "ANALISIS KOMPREHENSIF PROSES PEMBUKAAN CABANG (UPC) & SISTEM CERDAS SIMULASI"
    p_main.font.size = Pt(30)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE

    # Subtitle
    sub_box = slide1.shapes.add_textbox(Inches(1.0), Inches(3.9), Inches(11.0), Inches(1.2))
    tf_sub = sub_box.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = "Evaluasi Kinerja 1.039 Cabang (2024–2026), Rekonsiliasi Siklus Waktu, Uji Signifikansi Negosiator, Peta Geospasial 76 Wilayah, Vendor Risk, dan Prototipe REST API Auto-Routing."
    p_sub.font.size = Pt(14)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)

    # 4 Quick Metrics on Cover
    kpi_labels = [
        ("1.039 Cabang", "Total Evaluasi 2024-2026"),
        ("Rp 2,13 Miliar", "Penghematan Sewa 2026"),
        ("69,1 Hari", "Rata-rata Lead Time E2E"),
        ("15,60%", "Rerata Diskon Sewa 2026")
    ]
    for i, (val, lbl) in enumerate(kpi_labels):
        x = Inches(1.0 + i * 2.85)
        c = add_card(slide1, x, Inches(5.4), Inches(2.7), Inches(1.3), bg_color=NAVY_SECONDARY, border_color=BLUE_ACCENT)
        tb = slide1.shapes.add_textbox(x + Inches(0.1), Inches(5.5), Inches(2.5), Inches(1.1))
        tf = tb.text_frame
        p1 = tf.paragraphs[0]
        p1.text = val
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = WHITE
        p2 = tf.add_paragraph()
        p2.text = lbl
        p2.font.size = Pt(10)
        p2.font.color.rgb = RGBColor(148, 163, 184)

    # Footer
    f_box = slide1.shapes.add_textbox(Inches(1.0), Inches(6.9), Inches(11.3), Inches(0.4))
    tf_f = f_box.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "Penyusun: Data Analyst — Divisi Bisnis  |  Status: Dokumen Resmi Eksekutif  |  Tahun Anggaran 2026"
    p_f.font.size = Pt(10)
    p_f.font.color.rgb = RGBColor(148, 163, 184)

    # =========================================================================
    # SLIDE 2: RINGKASAN EKSEKUTIF & 4 PILAR TEMUAN UTAMA
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Ringkasan Eksekutif & 4 Pilar Utama Evaluasi Ekspansi")

    pillars = [
        ("1. Siklus Waktu & Lead Time", "Harmonisasi Paradoks Waktu", [
            "Lead time pasca-approval direksi stabil di 58,6 hari (~58 hari).",
            "Lead time end-to-end (termasuk survei) adalah 69,1 hari.",
            "Tahap berkas sewa terpangkas 66% (dari 32,7 hr ke 11,0 hr di 2026).",
            "Waktu tunggu Grand Opening terpangkas 60% (dari 14,0 hr ke 5,5 hr)."
        ], BLUE_ACCENT),
        ("2. Kinerja Tim Negosiasi", "Uji Signifikansi Bonita vs Mirza", [
            "Total saving 2026: Rp 2,13 Miliar (100% dari Bonita & Mirza).",
            "Bonita: 142 deal, saving Rp 1,40 Miliar (diskon 18,25%, durasi 8,5 hr).",
            "Mirza: 132 deal, saving Rp 728,4 Juta (diskon 14,05%, durasi 15,5 hr).",
            "Spesialisasi: Bonita tiket besar/alot, Mirza ekspansi & volume."
        ], GREEN_SUCCESS),
        ("3. Korelasi Durasi & Diskon", "The Law of Diminishing Returns", [
            "Uji korelasi Pearson r = -0.1040 membuktikan durasi lama TIDAK menambah diskon.",
            "Zona Emas (1–10 hari): 69,1% deal sukses meraih diskon 15–16%.",
            "Zona Bahaya (>14–20 hari): Tingkat gagal diskon melonjak hingga 27,5% (rugi waktu)."
        ], GOLD_ACCENT),
        ("4. Geospasial & Risiko Vendor", "Segmentasi Pasar & Kasus Bali", [
            "8 dari 10 wilayah paling alot berstatus Kabupaten (aset keluarga/warisan).",
            "Sarjono tetap Kontraktor Champion (SLA 89,0%, 73 unit).",
            "Trimo menanggung 92,3% proyek perintis Bali (12 unit); SLA terdistorsi kendala fery & banjar adat."
        ], RED_DANGER)
    ]

    for i, (title, sub, bullets, color) in enumerate(pillars):
        x = Inches(0.8 + i * 2.95)
        add_card(slide2, x, Inches(1.4), Inches(2.8), Inches(5.5))
        
        # Header banner inside card
        header_card = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.4), Inches(2.8), Inches(0.75))
        header_card.fill.solid()
        header_card.fill.fore_color.rgb = color
        header_card.line.fill.background()
        
        tb_h = slide2.shapes.add_textbox(x + Inches(0.1), Inches(1.42), Inches(2.6), Inches(0.7))
        tf_h = tb_h.text_frame
        tf_h.word_wrap = True
        p_h1 = tf_h.paragraphs[0]
        p_h1.text = title
        p_h1.font.size = Pt(12)
        p_h1.font.bold = True
        p_h1.font.color.rgb = WHITE
        p_h2 = tf_h.add_paragraph()
        p_h2.text = sub
        p_h2.font.size = Pt(9.5)
        p_h2.font.color.rgb = RGBColor(241, 245, 249)

        # Bullets
        tb_b = slide2.shapes.add_textbox(x + Inches(0.15), Inches(2.3), Inches(2.5), Inches(4.4))
        tf_b = tb_b.text_frame
        tf_b.word_wrap = True
        for b_idx, bullet in enumerate(bullets):
            p = tf_b.paragraphs[0] if b_idx == 0 else tf_b.add_paragraph()
            p.text = "• " + bullet
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(10)

    # =========================================================================
    # SLIDE 3: DEKOMPOSISI 7 SIKLUS WAKTU & HARMONISASI LEAD TIME
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Dekomposisi 7 Tahapan Siklus Waktu & Efisiensi Grand Opening")

    # Image Left
    img_path = os.path.join(CHART_DIR, "4_dekomposisi_siklus_waktu_upc.png")
    if os.path.exists(img_path):
        add_card(slide3, Inches(0.8), Inches(1.4), Inches(7.5), Inches(5.5))
        slide3.shapes.add_picture(img_path, Inches(0.9), Inches(1.5), width=Inches(7.3))

    # Right Card: Insights
    add_card(slide3, Inches(8.5), Inches(1.4), Inches(4.0), Inches(5.5))
    tb_c3 = slide3.shapes.add_textbox(Inches(8.7), Inches(1.6), Inches(3.6), Inches(5.1))
    tf_c3 = tb_c3.text_frame
    tf_c3.word_wrap = True

    p = tf_c3.paragraphs[0]
    p.text = "TEMUAN & REKONSILIASI KUNCI"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    bullets_s3 = [
        ("Harmonisasi Lead Time 58,6 vs 69,1 Hari", 
         "Lead time operasional pasca-approval direksi stabil di 58,6 hari (~58 hari). Sedangkan total proses end-to-end sejak pengajuan surveyor adalah 69,1 hari (selisih 10,5 hari adalah proses survei & approval awal)."),
        ("Akselerasi Pengumpulan Berkas (-66%)", 
         "Durasi pengumpulan berkas dari tgl negosiasi berakhir hingga ttd sewa terpangkas tajam dari 32,7 hari (2024) menjadi 11,0 hari di 2026."),
        ("Efisiensi Grand Opening Terpangkas 60%", 
         "Tahap tunggu GO berhasil dipotong dari 14,0 hari (2024) -> 9,3 hari (2025) -> kini hanya 5,5 hari (median 4 hari di 2026)."),
        ("Penyelarasan Tanggal Kalender Murni", 
         "Seluruh metrik waktu tunggu (seperti TTD sewa ke renovasi: 7,1 hari) dihitung murni hari kalender tanpa terpotong jam/detik pencatatan sistem.")
    ]
    for heading, body in bullets_s3:
        p_head = tf_c3.add_paragraph()
        p_head.text = f"\n📌 {heading}"
        p_head.font.size = Pt(11)
        p_head.font.bold = True
        p_head.font.color.rgb = BLUE_ACCENT
        p_body = tf_c3.add_paragraph()
        p_body.text = body
        p_body.font.size = Pt(9.5)
        p_body.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 4: EVALUASI KINERJA TIM NEGOSIASI 2026: BONITA VS MIRZA
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Evaluasi Kinerja Negosiator: The High-Value Specialist vs Speed Specialist")

    img_bonita_mirza = os.path.join(CHART_DIR, "2_komparasi_saving_bonita_mirza_2026.png")
    img_diskon_durasi = os.path.join(CHART_DIR, "3_komparasi_diskon_dan_durasi_2026.png")

    if os.path.exists(img_bonita_mirza):
        add_card(slide4, Inches(0.8), Inches(1.4), Inches(4.5), Inches(3.2))
        slide4.shapes.add_picture(img_bonita_mirza, Inches(0.9), Inches(1.45), width=Inches(4.3))

    if os.path.exists(img_diskon_durasi):
        add_card(slide4, Inches(0.8), Inches(4.8), Inches(4.5), Inches(2.3))
        slide4.shapes.add_picture(img_diskon_durasi, Inches(0.9), Inches(4.85), width=Inches(4.3))

    # Right Side: Comparison Table & Mann-Whitney Test Card
    add_card(slide4, Inches(5.6), Inches(1.4), Inches(6.9), Inches(5.7))
    tb_nego = slide4.shapes.add_textbox(Inches(5.8), Inches(1.55), Inches(6.5), Inches(5.4))
    tf_nego = tb_nego.text_frame
    tf_nego.word_wrap = True

    p = tf_nego.paragraphs[0]
    p.text = "KOMPARASI PERFORMA & UJI SIGNIFIKANSI STATISTIK"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    comparisons = [
        ("Bonita — The High-Value Negotiator",
         "• Menghasilkan penghematan Rp 1,40 Miliar (65,8% total saving nasional).\n• Rata-rata diskon 18,25% (median 18,7%) dengan tingkat sukses diskon 95,8% (142 deal).\n• Spesialisasi: Menangani ruko komersial bernilai sewa tinggi (>Rp 50 Jt) & negosiasi alot."),
        ("Mirza — The Speed & Volume Specialist",
         "• Menghasilkan penghematan Rp 728,4 Juta (34,2% total saving nasional).\n• Rata-rata diskon 14,05% (median 13,3%) dengan durasi rata-rata 15,5 hari (132 deal, puncak 28 deal/bln).\n• Spesialisasi: Volume deal bulanan tinggi & percepatan ekspansi cabang baru."),
        ("Hasil Uji Mann-Whitney U (p = 0.00012)",
         "• Uji statistik non-parametrik membuktikan keunggulan diskon Bonita signifikan secara ilmiah (p < 0.001) dan bukan kebetulan acak.\n• Rekomendasi: Penugasan berbasis Smart Routing (Bonita untuk tiket >Rp 50 Jt & alot, Mirza untuk speed & volume)."),
        ("Catatan Surveyor Lapangan (Dika & Salma)",
         "• 11 cabang pengajuan surveyor ditutup diskon 0% (Rp 0 saving) karena langsung deal tanpa lewat meja negosiasi resmi.")
    ]
    for head, desc in comparisons:
        p_h = tf_nego.add_paragraph()
        p_h.text = f"\n🔹 {head}"
        p_h.font.size = Pt(11)
        p_h.font.bold = True
        p_h.font.color.rgb = BLUE_ACCENT
        p_d = tf_nego.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(9.5)
        p_d.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 5: TREN BULANAN 2026: DINAMIKA VOLUME DEAL & PENGHEMATAN SEWA
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Tren Bulanan 2026: Dinamika Volume Deal & Penghematan Sewa (Bonita vs Mirza)")

    # Data for Tables
    bonita_data = [
        ["Bulan", "Deal", "Asking (Jt)", "Deal (Jt)", "Saving (Jt)", "Diskon", "Durasi", "Sukses"],
        ["Januari", "13", "458,0", "347,0", "111,0", "23,3%", "1,0 hr", "100%"],
        ["Februari", "16", "887,2", "712,0", "175,2", "18,1%", "11,8 hr", "100%"],
        ["Maret", "19", "897,0", "700,0", "197,0", "21,3%", "15,7 hr", "94,7%"],
        ["April", "26", "1.400,0", "1.181,0", "219,0", "16,7%", "11,7 hr", "92,3%"],
        ["Mei", "17", "1.022,5", "808,0", "214,5", "18,4%", "10,6 hr", "94,1%"],
        ["Juni", "23", "1.270,5", "1.029,0", "241,5", "16,4%", "5,9 hr", "95,7%"],
        ["Juli", "18", "971,0", "814,8", "156,2", "15,6%", "3,8 hr", "94,4%"],
        ["Agustus", "10", "495,0", "405,5", "89,5", "19,1%", "2,4 hr", "100%"],
        ["TOTAL '26", "142", "7.401,2", "5.997,3", "1.403,9", "18,3%", "8,5 hr", "95,8%"]
    ]

    mirza_data = [
        ["Bulan", "Deal", "Asking (Jt)", "Deal (Jt)", "Saving (Jt)", "Diskon", "Durasi", "Sukses"],
        ["Januari", "6", "152,0", "125,0", "27,0", "18,6%", "10,0 hr", "100%"],
        ["Februari", "17", "534,5", "448,7", "85,8", "17,3%", "11,1 hr", "100%"],
        ["Maret", "15", "475,0", "390,0", "85,0", "17,2%", "22,6 hr", "93,3%"],
        ["April", "28", "972,0", "825,8", "146,2", "13,8%", "17,3 hr", "92,9%"],
        ["Mei", "19", "835,0", "717,0", "118,0", "13,4%", "15,6 hr", "100%"],
        ["Juni", "19", "614,2", "543,0", "71,2", "10,2%", "8,5 hr", "84,2%"],
        ["Juli", "19", "971,0", "828,6", "142,4", "12,8%", "24,5 hr", "89,5%"],
        ["Agustus", "9", "416,0", "363,2", "52,8", "12,5%", "6,2 hr", "88,9%"],
        ["TOTAL '26", "132", "4.969,7", "4.241,3", "728,4", "14,0%", "15,5 hr", "93,2%"]
    ]

    # Helper function to populate PPT table
    def build_monthly_table(slide, x_pos, y_pos, width, height, data, header_color, total_bg_color, title_text):
        # Header banner badge
        badge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_pos, y_pos, width, Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = header_color
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = title_text
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = WHITE
        p_b.alignment = PP_ALIGN.CENTER

        # Table Shape
        rows = len(data)
        cols = len(data[0])
        tbl_shape = slide.shapes.add_table(rows, cols, x_pos, y_pos + Inches(0.38), width, height - Inches(0.38))
        tbl = tbl_shape.table

        # Set specific column widths
        col_w = [Inches(0.95), Inches(0.45), Inches(0.80), Inches(0.80), Inches(0.80), Inches(0.65), Inches(0.65), Inches(0.65)]
        for c_idx, w in enumerate(col_w):
            tbl.columns[c_idx].width = w

        for r_idx, row in enumerate(data):
            is_header = (r_idx == 0)
            is_total = (r_idx == rows - 1)
            for c_idx, val in enumerate(row):
                cell = tbl.cell(r_idx, c_idx)
                cell.text = str(val)
                p = cell.text_frame.paragraphs[0]
                p.font.name = "Calibri"
                p.font.size = Pt(8)
                
                # Alignment
                if c_idx == 0:
                    p.alignment = PP_ALIGN.LEFT
                else:
                    p.alignment = PP_ALIGN.RIGHT if c_idx in [1, 2, 3, 4] else PP_ALIGN.CENTER
                
                # Colors
                cell.fill.solid()
                if is_header:
                    cell.fill.fore_color.rgb = NAVY_PRIMARY
                    p.font.bold = True
                    p.font.color.rgb = WHITE
                    p.font.size = Pt(8)
                elif is_total:
                    cell.fill.fore_color.rgb = total_bg_color
                    p.font.bold = True
                    p.font.color.rgb = NAVY_PRIMARY
                    p.font.size = Pt(8.5)
                else:
                    bg_c = WHITE if r_idx % 2 == 1 else RGBColor(241, 245, 249)
                    cell.fill.fore_color.rgb = bg_c
                    p.font.color.rgb = TEXT_DARK

    # Render Bonita Table (Left)
    build_monthly_table(
        slide5, 
        Inches(0.8), Inches(1.35), Inches(5.75), Inches(4.1), 
        bonita_data, 
        BLUE_ACCENT, 
        RGBColor(219, 234, 254), 
        "TIM BONITA: 142 Deal | Saving Rp 1,40 M (65,8%) | Diskon 18,25%"
    )

    # Render Mirza Table (Right)
    build_monthly_table(
        slide5, 
        Inches(6.75), Inches(1.35), Inches(5.75), Inches(4.1), 
        mirza_data, 
        GREEN_SUCCESS, 
        RGBColor(209, 250, 229), 
        "TIM MIRZA: 132 Deal | Saving Rp 728,4 Jt (34,2%) | Diskon 14,05%"
    )

    # Bottom Takeaway Card
    add_card(slide5, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.5), bg_color=WHITE)
    tb_bot = slide5.shapes.add_textbox(Inches(0.95), Inches(5.65), Inches(11.4), Inches(1.4))
    tf_bot = tb_bot.text_frame
    tf_bot.word_wrap = True

    p_t = tf_bot.paragraphs[0]
    p_t.text = "💡 INSIGHT & KOMPARASI STRATEGIS BULANAN 2026"
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = NAVY_PRIMARY

    bullets_trend = [
        ("Konsistensi Penghematan Sewa", "Bonita membukukan rata-rata saving Rp 175,5 Jt/bln (puncak Juni Rp 241,5 Jt). Mirza membukukan rata-rata saving Rp 91,1 Jt/bln (puncak April Rp 146,2 Jt). Total kedua tim mengamankan Rp 2,13 Miliar."),
        ("Akselerasi Kecepatan Closing", "Bonita mencatatkan lonjakan efisiensi durasi pada semester II (Juli 3,8 hr, Agustus 2,4 hr). Mirza mempertahankan peran kunci volume tinggi (puncak April 28 deal) untuk mengejar target ekspansi."),
        ("Tingkat Keberhasilan Diskon", "Bonita membukukan success rate 95,8% (hanya 6 deal tanpa diskon dari 142 deal), Mirza membukukan 93,2% (9 deal tanpa diskon dari 132 deal). Keduanya memiliki kepatuhan sangat tinggi terhadap mandat efisiensi sewa.")
    ]
    for h_txt, b_txt in bullets_trend:
        p_b = tf_bot.add_paragraph()
        p_b.text = f"• {h_txt}: {b_txt}"
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 6: THE LAW OF DIMINISHING RETURNS (DURASI VS DISKON)
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Hukum Penurunan Hasil (Law of Diminishing Returns) pada Durasi Negosiasi")

    img_korelasi = os.path.join(CHART_DIR, "7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png")
    if os.path.exists(img_korelasi):
        add_card(slide6, Inches(0.8), Inches(1.4), Inches(7.5), Inches(5.5))
        slide6.shapes.add_picture(img_korelasi, Inches(0.9), Inches(1.5), width=Inches(7.3))

    # Right Card: Insights
    add_card(slide6, Inches(8.5), Inches(1.4), Inches(4.0), Inches(5.5))
    tb_c6 = slide6.shapes.add_textbox(Inches(8.7), Inches(1.6), Inches(3.6), Inches(5.1))
    tf_c6 = tb_c6.text_frame
    tf_c6.word_wrap = True

    p = tf_c6.paragraphs[0]
    p.text = "FAKTA EMPIRIS & ANALISIS BIVARIAT"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    bullets_s6 = [
        ("Korelasi Pearson r = -0.1040 (p = 0.085)",
         "Hubungan durasi tawar-menawar terhadap diskon bernilai negatif. Menambah waktu negosiasi secara empiris TIDAK meningkatkan diskon sewa (R² hanya 1,1%)."),
        ("Zona Emas Negosiasi (1–10 Hari)",
         "Menyerap 69,1% transaksi (190 deal). Rata-rata diskon mencapai 15%–16% dengan tingkat keberhasilan diskon tertinggi 93%–97%."),
        ("Zona Bahaya (>14–20 Hari)",
         "Tingkat kegagalan meraih diskon (0% saving) melonjak drastis hingga 27,5% pada durasi di atas 20 hari. Biaya tunda (cost of delay) jauh melampaui potensi diskon tambahan."),
        ("Rekomendasi Kebijakan SLA",
         "Terapkan batas cut-off tegas: Evaluasi eskalasi pada Hari ke-10, dan cut-off maksimal Hari ke-14.")
    ]
    for heading, body in bullets_s6:
        p_head = tf_c6.add_paragraph()
        p_head.text = f"\n⚠️ {heading}"
        p_head.font.size = Pt(11)
        p_head.font.bold = True
        p_head.font.color.rgb = RED_DANGER if "Bahaya" in heading or "r = -0.1040" in heading else BLUE_ACCENT
        p_body = tf_c6.add_paragraph()
        p_body.text = body
        p_body.font.size = Pt(9.5)
        p_body.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 7: PETA GEOSPASIAL & KARAKTERISTIK PASAR WILAYAH
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Indeks Geospasial Kemudahan Negosiasi: Pemetaan 76 Kabupaten/Kota")

    img_geo1 = os.path.join(CHART_DIR, "5_peta_geospasial_kemudahan_nego_2026.png")
    if os.path.exists(img_geo1):
        add_card(slide7, Inches(0.8), Inches(1.4), Inches(5.8), Inches(5.5))
        slide7.shapes.add_picture(img_geo1, Inches(0.9), Inches(1.5), width=Inches(5.6))

    # Right Card: Geospatial Insights
    add_card(slide7, Inches(6.8), Inches(1.4), Inches(5.7), Inches(5.5))
    tb_geo = slide7.shapes.add_textbox(Inches(7.0), Inches(1.55), Inches(5.3), Inches(5.2))
    tf_geo = tb_geo.text_frame
    tf_geo.word_wrap = True

    p = tf_geo.paragraphs[0]
    p.text = "SEGMENTASI WILAYAH & POLA KULTURAL PROPERTI"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    geo_bullets = [
        ("Zona Hijau Zamrud (Sangat Mudah, Skor >= 70)",
         "• Wilayah: Cilacap (92,2), Tasikmalaya (76,7), Serang (75,2), Jakarta Pusat (73,0).\n• Karakteristik: Pemilik kooperatif, durasi closing rata-rata 3–6 hari, diskon rata-rata >18%."),
        ("Zona Merah Tua (Sangat Alot, Skor < 40)",
         "• Wilayah: Bandung Barat (19,0), Grobogan (20,0), Cirebon Kota (28,3), Cianjur (28,5).\n• Karakteristik: Rata-rata durasi membengkak >20 hari, diskon rendah (< 5%–10%)."),
        ("Temuan Kultural: Dikotomi Kabupaten vs Kota",
         "• 8 dari 10 wilayah paling alot di Indonesia berstatus KABUPATEN.\n• Faktor Penyebab: Ruko di kabupaten didominasi aset warisan keluarga majemuk yang memerlukan konsensus banyak ahli waris.\n• Anomali Kota: Kota Cirebon & Surakarta alot karena dihuni saudagar tua dengan modal kas kuat."),
        ("Penerapan TopoJSON BPS",
         "• Seluruh 76 kab/kota aktif 2026 terpetakan 100% tanpa ada daerah yang lolos (zero unmapped).")
    ]
    for h, b in geo_bullets:
        p_h = tf_geo.add_paragraph()
        p_h.text = f"\n📍 {h}"
        p_h.font.size = Pt(10.5)
        p_h.font.bold = True
        p_h.font.color.rgb = BLUE_ACCENT
        p_b = tf_geo.add_paragraph()
        p_b.text = b
        p_b.font.size = Pt(9)
        p_b.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 8: VENDOR RISK QUADRANT & KASUS KHUSUS BALI
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Vendor Risk Quadrant: Evaluasi Kontraktor & Analisis Beban Khusus Bali")

    img_vendor_q = os.path.join(CHART_DIR, "8_peta_risiko_vendor_kuadran.png")
    if os.path.exists(img_vendor_q):
        add_card(slide8, Inches(0.8), Inches(1.4), Inches(6.8), Inches(5.5))
        slide8.shapes.add_picture(img_vendor_q, Inches(0.9), Inches(1.5), width=Inches(6.6))

    # Right Card: Vendor Insights
    add_card(slide8, Inches(7.8), Inches(1.4), Inches(4.7), Inches(5.5))
    tb_v = slide8.shapes.add_textbox(Inches(8.0), Inches(1.55), Inches(4.3), Inches(5.2))
    tf_v = tb_v.text_frame
    tf_v.word_wrap = True

    p = tf_v.paragraphs[0]
    p.text = "ANALISIS PERFORMA KONTRAKTOR & BALI"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY

    v_bullets = [
        ("Sarjono — Champion Kuadran I",
         "• Volume 73 unit, Kepatuhan SLA tertinggi 89,0%, durasi rata-rata tercepat 14,9 hari.\n• Mitra paling andal untuk pengerjaan cabang prioritas."),
        ("CV Rizki Mitra Abadi — Turnaround",
         "• Lonjakan kepatuhan SLA dari 44,6% (2025) ke 66,7% (2026), mendekati Kuadran I."),
        ("Dekomposisi Realitas Trimo (SLA 15,1%)",
         "• Menanggung 92,3% proyek perintis Bali (12 dari 13 cabang) yang dihindari kontraktor lain.\n• Proyek Bali: SLA 0% (rata-rata 43,9 hari) akibat antrean fery Ketapang-Gilimanuk & upacara Banjar adat.\n• Proyek Reguler Jawa: SLA 19,5% (31,8 hari).\n• Evaluasi objektif: Kinerja Trimo terdistorsi beban perintis pulau luar."),
        ("Mitra Dihentikan / Underperformer",
         "• CV Cahaya Kemakmuran dinonaktifkan (SLA '25 hanya 15,4%).\n• Dimas Andri Sulistyo (mitra baru) masuk zona pengawasan ketat (SLA 11,1%).")
    ]
    for h, b in v_bullets:
        p_h = tf_v.add_paragraph()
        p_h.text = f"\n🔨 {h}"
        p_h.font.size = Pt(10)
        p_h.font.bold = True
        p_h.font.color.rgb = GREEN_SUCCESS if "Sarjono" in h else (GOLD_ACCENT if "Trimo" in h else BLUE_ACCENT)
        p_b = tf_v.add_paragraph()
        p_b.text = b
        p_b.font.size = Pt(8.8)
        p_b.font.color.rgb = TEXT_DARK

    # =========================================================================
    # SLIDE 9: INOVASI SISTEM CERDAS & AUTO-ROUTING
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Inovasi Sistem Cerdas: Auto-Routing Negosiasi & Dashboard Interaktif")

    cards_feat = [
        ("Sistem Rekomendasi Auto-Routing",
         "Secara otomatis mencocokkan profil ruko dengan kapabilitas negosiator:\n\n"
         "• Strict Fair Workload Guardrail: Selisih beban proyek aktif Bonita vs Mirza dibatasi MAKSIMAL 2 PROYEK (mencegah ketimpangan & kecemburuan tim).\n"
         "• Ruko Tiket Besar (>Rp 50 Jt/thn) -> Dialokasikan ke BONITA (Maksimalisasi Saving Diskon).\n"
         "• Ruko Cepat / Kebutuhan Akselerasi GO -> Dialokasikan ke MIRZA (Kecepatan Closing).\n"
         "• Mencegah deal 0% akibat surveyor langsung deal.",
         BLUE_ACCENT),
        ("Mesin Prediksi Multi-Variabel",
         "Menghitung estimasi realistik sebelum negosiasi dimulai:\n\n"
         "• Estimasi SLA Durasi Nego (Median historis wilayah).\n"
         "• Target Diskon Rasional (Batas atas & bawah).\n"
         "• Indeks Kemudahan Wilayah (Skor 0–100).\n"
         "• Wilayah Acuan Pembanding Terdekat (Clustering Euclidean Space 3 Dimensi).",
         GREEN_SUCCESS),
        ("REST API & Web Dashboard Siap Pakai",
         "Telah dilengkapi antarmuka modern yang siap diintegrasikan:\n\n"
         "• Endpoint JSON REST API (/api/routing, /api/predict).\n"
         "• Dashboard Web Interaktif (HTML5 + Tailwind CSS).\n"
         "• Cloudflare Global Tunnel (Akses aman dari laptop/HP atasan tanpa login rumit).",
         GOLD_ACCENT)
    ]

    for i, (f_title, f_desc, f_color) in enumerate(cards_feat):
        x = Inches(0.8 + i * 3.95)
        add_card(slide9, x, Inches(1.4), Inches(3.8), Inches(5.5))
        
        # Banner
        b_box = slide9.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.4), Inches(3.8), Inches(0.8))
        b_box.fill.solid()
        b_box.fill.fore_color.rgb = f_color
        b_box.line.fill.background()
        
        tb = slide9.shapes.add_textbox(x + Inches(0.15), Inches(1.45), Inches(3.5), Inches(0.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f_title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = WHITE
        
        tb_body = slide9.shapes.add_textbox(x + Inches(0.2), Inches(2.3), Inches(3.4), Inches(4.4))
        tf_body = tb_body.text_frame
        tf_body.word_wrap = True
        p_body = tf_body.paragraphs[0]
        p_body.text = f_desc
        p_body.font.size = Pt(10.5)
        p_body.font.color.rgb = TEXT_DARK
        p_body.space_after = Pt(6)

    # =========================================================================
    # SLIDE 10: 5 REKOMENDASI MANAJERIAL STRATEGIS DIREKSI
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)
    add_header(slide10, "5 Rekomendasi Manajerial Strategis untuk Direksi & Manajemen PGI")

    recs = [
        ("1. Kebijakan Cut-off Negosiasi 14 Hari Kerja",
         "Hentikan tawar-menawar yang melampaui 14 hari kerja. Analisis membuktikan durasi >14 hari menghasilkan kegagalan diskon hingga 27,5% dan menimbulkan biaya sewa terbuang (cost of delay)."),
        ("2. Implementasi Sistem Cerdas Auto-Routing",
         "Tugaskan Bonita khusus untuk ruko berharga sewa di atas Rp 50 Juta atau wilayah alot, dan tugaskan Mirza untuk target ekspansi cepat dan ruko komersial standar."),
        ("3. Sentralisasi 100% Negosiasi di Tim Resmi",
         "Hentikan penerimaan deal sewa langsung oleh surveyor lapangan (kasus 11 cabang diskon 0% di 2026). Seluruh penawaran wajib melewati verifikasi Tim Negosiasi."),
        ("4. Diferensiasi SLA & Insentif Kontraktor Luar Pulau (Kasus Bali)",
         "Berikan SLA realistis (45 hari) untuk ekspansi Pulau Bali dan wilayah kepulauan lainnya, serta evaluasi alokasi beban berimbang agar tidak membebani satu kontraktor tunggal (Trimo)."),
        ("5. Integrasi Dashboard Pemantauan Real-Time",
         "Manfaatkan Web Dashboard & REST API yang telah dibangun sebagai Early Warning System pemantauan harian proses pembukaan cabang.")
    ]

    for i, (r_title, r_desc) in enumerate(recs):
        y = Inches(1.4 + i * 1.08)
        card_r = add_card(slide10, Inches(0.8), y, Inches(11.7), Inches(0.95), bg_color=WHITE)
        
        # Left tag icon
        icon_box = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), y, Inches(0.2), Inches(0.95))
        icon_box.fill.solid()
        icon_box.fill.fore_color.rgb = BLUE_ACCENT if i % 2 == 0 else GOLD_ACCENT
        icon_box.line.fill.background()
        
        tb = slide10.shapes.add_textbox(Inches(1.15), y + Inches(0.08), Inches(11.2), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.text = r_title
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_PRIMARY
        
        p2 = tf.add_paragraph()
        p2.text = r_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = TEXT_DARK

    prs.save(OUTPUT_PPTX)
    print(f"[✓] Presentasi PowerPoint berhasil dibuat: {OUTPUT_PPTX}")
    print(f"    Ukuran file: {os.path.getsize(OUTPUT_PPTX) / 1024:.1f} KB, Total Slides: {len(prs.slides)}")

if __name__ == "__main__":
    create_deck()
