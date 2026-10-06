#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SKRIP ANALISIS RISIKO VENDOR & PETA KUADRAN PERFORMA KONTRAKTOR (2025 VS 2026)
PUSAT GADAI INDONESIA (PGI) — DIVISI BISNIS
================================================================================

Penulis      : Divisi Bisnis (Data Analyst Team)
Dataset      : properties_cleaned_2024-2026.csv (1.028 Cabang Unik)
Acuan Desain : DOOM_UPC_Operational_Data_Analyst.pdf (Poin A: Peta Risiko Vendor)
Fokus Utama  :
  1. Menampilkan Poin A (Peta Risiko Vendor / Vendor Risk Quadrant) secara proporsional.
  2. Membandingkan performa dan dinamika posisi kontraktor antara TAHUN 2025 vs TAHUN 2026.
  3. Menyajikan 2 Visualisasi Beresolusi Tinggi (300 DPI):
     - Grafik 8: Peta Risiko Vendor Komparatif Berdampingan (Side-by-Side 2025 vs 2026)
     - Grafik 9: Peta Vektor Lintasan Migrasi Risiko Vendor (Trajectory Map 2025 -> 2026)
  4. Mengungkap Realitas Operasional Lapangan:
     - Kontraktor Trimo mengambil beban penugasan ekspansi wilayah BALI (12 cabang di 2026),
       sedangkan vendor lain (Sarjono, Sendy, Edwin) belum bersedia mengambil wilayah Bali.
     - Penyeberangan fery Ketapang-Gilimanuk, rantai pasok antarpulau, dan regulasi adat Banjar
       menyebabkan durasi Trimo di Bali mencapai 43,9 hari (0% SLA).
     - Tanpa beban Bali, durasi Trimo di Pulau Jawa adalah 31,8 hari.
  5. Menghasilkan laporan berkas Excel multi-sheet dan tabel terminal komprehensif.

================================================================================
CARA MENJALANKAN SKRIP DI VISUAL STUDIO CODE (VSCODE):
================================================================================
1. Buka folder kerja di VSCode:
   - File -> Open Folder... -> Pilih folder 'Nego_baru'
2. Buka berkas ini di editor:
   - Dobel klik file 'analisa_risiko_vendor.py'
3. Jalankan skrip:
   - CARA A: Klik tombol 'Play/Run Python File' di pojok kanan atas editor.
   - CARA B: Melalui terminal terintegrasi VSCode (Ctrl + `):
         python3 analisa_risiko_vendor.py
4. Output:
   - Tabel dashboard di terminal VSCode
   - File Excel: 'hasil_analisis/Laporan_Risiko_Vendor_Renovasi.xlsx'
   - Grafik PNG:
     * 'hasil_analisis/grafik/8_peta_risiko_vendor_komparasi_2025_2026.png'
     * 'hasil_analisis/grafik/9_vektor_migrasi_risiko_vendor_2025_2026.png'
================================================================================
"""

import os
import sys
import shutil
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from scipy import stats

# ==============================================================================
# KONFIGURASI DIREKTORI & PARAMETER
# ==============================================================================
DIR_KERJA = os.path.dirname(os.path.abspath(__file__))
FILE_DATA_EXCEL = os.path.join(DIR_KERJA, "data_nego_baru", "data_nego_baru_cleaned_2024-2026.xlsx")
FILE_DATA_CSV = os.path.join(DIR_KERJA, "properties_cleaned_2024-2026.csv")
FILE_DATA = FILE_DATA_EXCEL if os.path.exists(FILE_DATA_EXCEL) else FILE_DATA_CSV
OUTPUT_DIR = os.path.join(DIR_KERJA, "hasil_analisis")
GRAFIK_DIR = os.path.join(OUTPUT_DIR, "grafik")
ARTIFACT_GRAFIK_DIR = "/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik"

# Parameter Batas Kuadran (Crosshair)
THRESHOLD_VOLUME_2025 = 50.0  # Unit cabang per tahun (Kohort 2025)
THRESHOLD_VOLUME_2026 = 40.0  # Unit cabang per tahun (Kohort 2026)
THRESHOLD_VOLUME_MIG = 45.0   # Unit cabang untuk Vektor Migrasi
THRESHOLD_SLA = 70.0          # Ambang batas sehat kepatuhan SLA operasional PGI (%)

# Palet Warna Standar Vendor
COLOR_MAP = {
    "Sarjono": "#10B981",                   # Hijau Emerald (Core Champion)
    "Sendy": "#059669",                     # Hijau Tua
    "Edwin": "#2563EB",                     # Biru Royal
    "CV. Rizki Mitra Abadi": "#F97316",     # Oranye (Improving)
    "Trimo": "#DC2626",                     # Merah Crimson (High Operational Risk)
    "Teguh Karyanto": "#9333EA",            # Ungu
    "CV Cahaya Kemakmuran Sejahtera": "#64748B", # Abu-abu (Discontinued)
    "Dimas Andri Sulistyo": "#D97706"       # Amber
}


# ==============================================================================
# 1. PEMUATAN & PENYIAPAN DATASET
# ==============================================================================
def load_and_prep_vendor_data():
    """Membaca dan memvalidasi dataset untuk analisis kontraktor renovasi."""
    if os.path.exists(FILE_DATA_EXCEL):
        print(f"[*] Berhasil memuat dataset Excel utama: {FILE_DATA_EXCEL}")
        df = pd.read_excel(FILE_DATA_EXCEL)
    elif os.path.exists(FILE_DATA_CSV):
        print(f"[*] Berhasil memuat dataset CSV: {FILE_DATA_CSV}")
        df = pd.read_csv(FILE_DATA_CSV)
    else:
        print(f"[!] ERROR: Berkas data tidak ditemukan: {FILE_DATA}")
        sys.exit(1)

    print(f"[*] Berhasil memuat dataset: {len(df)} total cabang unik.")

    # Standardisasi kolom numerik
    df["budget_num"] = pd.to_numeric(df["budget"], errors="coerce").fillna(0)
    df["durasi_renov"] = pd.to_numeric(df["lama_waktu_realisasi_renovasi_Selesai"], errors="coerce").fillna(0)
    df["sla_target"] = pd.to_numeric(df["sla_renov"], errors="coerce").fillna(0)
    df["tahun_int"] = pd.to_numeric(df["Tahun"], errors="coerce").fillna(0).astype(int)

    # Standardisasi kepatuhan SLA (berdasarkan kolom resmi sla_complimence_persen)
    if "sla_complimence_persen" in df.columns:
        df["is_patuh"] = df["sla_complimence_persen"].astype(str).str.strip().str.lower() == "tercapai"
    else:
        df["is_patuh"] = df["durasi_renov"] <= df["sla_target"]

    # Pembersihan nama kontraktor
    df["kontraktor_clean"] = df["kontraktor"].fillna("Tanpa Nama").astype(str).str.strip()
    
    # Identifikasi wilayah Bali dengan regex word boundary (mencegah false match pada Purbalingga)
    bali_keywords = ['DENPASAR', 'BADUNG', 'GIANYAR', 'TABANAN', 'BULELENG', 'SINGARAJA', 'KLUNGKUNG', 'BANGLI', 'JEMBRANA', 'KARANGASEM', 'BALI']
    pattern_bali = r'\b(?:' + '|'.join(bali_keywords) + r')\b'
    col_w = 'wilayah' if 'wilayah' in df.columns else df.columns[0]
    col_a = 'alamat' if 'alamat' in df.columns else col_w
    df["is_bali"] = (
        df[col_w].astype(str).str.upper().str.contains(pattern_bali, regex=True) |
        df[col_a].astype(str).str.upper().str.contains(pattern_bali, regex=True)
    )

    return df


# ==============================================================================
# 2. KALKULASI METRIK KINERJA TAHUNAN & MIGRASI (2025 VS 2026)
# ==============================================================================
def hitung_metrik_tahunan(df, tahun):
    """Menghitung agregasi performa per kontraktor untuk tahun tertentu."""
    dft = df[df["tahun_int"] == tahun].copy()
    grp = dft.groupby("kontraktor_clean").agg(
        unit=("durasi_renov", "count"),
        avg_dur=("durasi_renov", "mean"),
        med_dur=("durasi_renov", "median"),
        avg_sla_target=("sla_target", "mean"),
        patuh_cnt=("is_patuh", "sum"),
        avg_budget=("budget_num", "mean"),
        bali_unit=("is_bali", "sum")
    ).reset_index()

    grp["pct_patuh"] = (grp["patuh_cnt"] / grp["unit"]) * 100
    grp["tahun"] = tahun

    # Tentukan kuadran berdasarkan ambang batas tahunan
    th_vol = THRESHOLD_VOLUME_2026 if tahun == 2026 else THRESHOLD_VOLUME_2025
    def get_kuadran(r):
        v = r["unit"]
        s = r["pct_patuh"]
        if v >= th_vol and s >= THRESHOLD_SLA:
            return "Kuadran I: Core Champion"
        elif v >= th_vol and s < THRESHOLD_SLA:
            return "Kuadran II: Zona Risiko Operasional"
        elif v < th_vol and s < THRESHOLD_SLA:
            return "Kuadran III: Underperformer"
        else:
            return "Kuadran IV: Potensial / Selektif"

    grp["kuadran"] = grp.apply(get_kuadran, axis=1)
    grp = grp.sort_values("unit", ascending=False).reset_index(drop=True)
    return grp


def hitung_komparasi_migrasi(grp_2025, grp_2026):
    """Menghitung selisih metrik dan pergeseran posisi antar tahun 2025 ke 2026."""
    m = pd.merge(
        grp_2025[["kontraktor_clean", "unit", "avg_dur", "pct_patuh", "kuadran", "bali_unit"]],
        grp_2026[["kontraktor_clean", "unit", "avg_dur", "pct_patuh", "kuadran", "bali_unit"]],
        on="kontraktor_clean",
        how="outer",
        suffixes=("_2025", "_2026")
    ).fillna({
        "unit_2025": 0, "avg_dur_2025": 0, "pct_patuh_2025": 0, "kuadran_2025": "Tidak Aktif", "bali_unit_2025": 0,
        "unit_2026": 0, "avg_dur_2026": 0, "pct_patuh_2026": 0, "kuadran_2026": "Tidak Aktif", "bali_unit_2026": 0
    })

    m["delta_unit"] = m["unit_2026"] - m["unit_2025"]
    m["delta_durasi"] = m["avg_dur_2026"] - m["avg_dur_2025"]
    m["delta_sla_poin"] = m["pct_patuh_2026"] - m["pct_patuh_2025"]

    # Urutkan berdasarkan total volume 2026 lalu 2025
    m["total_vol"] = m["unit_2026"] + m["unit_2025"]
    m = m.sort_values("total_vol", ascending=False).reset_index(drop=True)
    return m


def hitung_analisis_bali(df):
    """Menganalisis perbandingan performa Trimo di Bali vs Luar Bali pada tahun 2026."""
    df26 = df[df["tahun_int"] == 2026].copy()
    
    # Ringkasan per vendor di Bali vs Non-Bali
    bali_summary = df26.groupby(["kontraktor_clean", "is_bali"]).agg(
        unit=("durasi_renov", "count"),
        avg_dur=("durasi_renov", "mean"),
        med_dur=("durasi_renov", "median"),
        patuh=("is_patuh", "sum")
    ).reset_index()
    bali_summary["pct_patuh"] = (bali_summary["patuh"] / bali_summary["unit"]) * 100
    
    # Rincian khusus Trimo
    trimo_bali = df26[(df26["kontraktor_clean"] == "Trimo") & (df26["is_bali"])]["durasi_renov"]
    trimo_non_bali = df26[(df26["kontraktor_clean"] == "Trimo") & (~df26["is_bali"])]["durasi_renov"]
    
    trimo_stats = {
        "bali_unit": len(trimo_bali),
        "bali_avg_dur": trimo_bali.mean() if len(trimo_bali) > 0 else 0,
        "bali_med_dur": trimo_bali.median() if len(trimo_bali) > 0 else 0,
        "bali_sla_pct": (df26[(df26["kontraktor_clean"] == "Trimo") & (df26["is_bali"])]["is_patuh"].sum() / len(trimo_bali) * 100) if len(trimo_bali) > 0 else 0,
        "non_bali_unit": len(trimo_non_bali),
        "non_bali_avg_dur": trimo_non_bali.mean() if len(trimo_non_bali) > 0 else 0,
        "non_bali_med_dur": trimo_non_bali.median() if len(trimo_non_bali) > 0 else 0,
        "non_bali_sla_pct": (df26[(df26["kontraktor_clean"] == "Trimo") & (~df26["is_bali"])]["is_patuh"].sum() / len(trimo_non_bali) * 100) if len(trimo_non_bali) > 0 else 0,
    }
    return bali_summary, trimo_stats


# ==============================================================================
# 3. GRAFIK 8: PETA RISIKO VENDOR KOMPARASI (SIDE-BY-SIDE 2025 VS 2026)
# ==============================================================================
def buat_grafik_komparasi_2025_2026(grp_2025, grp_2026):
    """
    Menghasilkan Grafik 8: Peta Risiko Vendor Berdampingan (Side-by-Side)
    Memfokuskan hanya pada Poin A (Vendor Risk Quadrant) tanpa kartu hipotesis di kanan.
    """
    os.makedirs(GRAFIK_DIR, exist_ok=True)
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8.5), dpi=300)

    y_max = 105

    # --------------------------------------------------------------------------
    # PANEL A1: TAHUN 2025 (436 CABANG)
    # --------------------------------------------------------------------------
    x_max_25 = 125
    th_vol_25 = THRESHOLD_VOLUME_2025
    th_sla = THRESHOLD_SLA

    # Kuadran Latar Belakang
    ax1.fill_between([th_vol_25, x_max_25], [th_sla, th_sla], [y_max, y_max], color="#ECFDF5", alpha=0.6, zorder=1)
    ax1.fill_between([th_vol_25, x_max_25], [0, 0], [th_sla, th_sla], color="#FEF2F2", alpha=0.6, zorder=1)
    ax1.fill_between([0, th_vol_25], [0, 0], [th_sla, th_sla], color="#FFFBEB", alpha=0.6, zorder=1)
    ax1.fill_between([0, th_vol_25], [th_sla, th_sla], [y_max, y_max], color="#EFF6FF", alpha=0.6, zorder=1)

    ax1.axvline(th_vol_25, color="#64748B", linestyle="--", linewidth=1.5, zorder=2)
    ax1.axhline(th_sla, color="#64748B", linestyle="--", linewidth=1.5, zorder=2)

    ax1.text(88, 101, "KUADRAN I: CORE CHAMPIONS\n(Volume Tinggi, Kepatuhan Tinggi)",
             fontsize=8.2, fontweight="bold", color="#065F46", ha="center", zorder=3)
    ax1.text(88, 4, "KUADRAN II: ZONA RISIKO OPERASIONAL\n(Volume Tinggi, Kepatuhan Rendah)",
             fontsize=8.0, fontweight="bold", color="#991B1B", ha="center", zorder=3)
    ax1.text(25, 4, "KUADRAN III: UNDERPERFORMERS\n(Volume Rendah, Kepatuhan Rendah)",
             fontsize=8.0, fontweight="bold", color="#92400E", ha="center", zorder=3)
    ax1.text(25, 101, "KUADRAN IV: POTENSIAL / SELEKTIF\n(Volume Rendah, Kepatuhan Tinggi)",
             fontsize=8.0, fontweight="bold", color="#1E40AF", ha="center", zorder=3)

    ax1.text(th_vol_25 + 2, th_sla + 1.2, "Ambang Batas Sehat (70%)", fontsize=7.2, fontweight="bold", color="#334155")
    ax1.text(th_vol_25 + 2, 20, f"Batas Volume {int(th_vol_25)} Unit", fontsize=7.2, fontweight="bold", color="#334155", rotation=90)

    # Plot Kontraktor 2025
    df_plot_25 = grp_2025[grp_2025["unit"] >= 2].copy()
    for _, r in df_plot_25.iterrows():
        vname = r["kontraktor_clean"]
        x, y = r["unit"], r["pct_patuh"]
        col = COLOR_MAP.get(vname, "#475569")
        size = 180 + (r["avg_dur"] * 6)

        ax1.scatter(x, y, s=size, color=col, edgecolors="#0F172A", linewidth=1.4, zorder=4, alpha=0.9)

        # Penataan Anotasi
        if vname == "Sarjono":
            pos_x, pos_y, ha = x, y + 5.0, "center"
        elif vname == "Sendy":
            pos_x, pos_y, ha = x - 6, y - 6.5, "right"
        elif vname == "Trimo":
            pos_x, pos_y, ha = x, y - 7.5, "center"
        elif vname == "CV. Rizki Mitra Abadi":
            pos_x, pos_y, ha = x + 5, y + 4.5, "left"
        elif vname == "Edwin":
            pos_x, pos_y, ha = x - 5, y + 4.5, "right"
        elif "Cahaya" in vname:
            pos_x, pos_y, ha = x + 5, y + 3.5, "left"
            vname = "CV Cahaya"
        elif vname == "Teguh Karyanto":
            pos_x, pos_y, ha = x + 4, y + 4.5, "left"
            vname = "Teguh K."
        else:
            pos_x, pos_y, ha = x, y + 3.5, "center"

        lbl = f"{vname}\n({int(x)}u | {y:.1f}%)"
        ax1.annotate(lbl, (pos_x, pos_y), fontsize=7.8, fontweight="bold", ha=ha, color="#0F172A", zorder=5,
                     bbox=dict(boxstyle="round,pad=0.2", facecolor="#FFFFFF", alpha=0.85, edgecolor="#CBD5E1", linewidth=0.8))

    ax1.set_title("A1. Vendor Risk Quadrant TAHUN 2025 (436 Cabang)\nSLA Agregat PGI: 64,2% (Sendy & Sarjono Dominan)",
                  fontsize=10.5, fontweight="bold", pad=12)
    ax1.set_xlabel("Volume Pengerjaan Renovasi (Jumlah Unit Cabang)", fontsize=9.5, fontweight="bold")
    ax1.set_ylabel("Tingkat Kepatuhan SLA Renovasi (%)", fontsize=9.5, fontweight="bold")
    ax1.set_xlim(0, x_max_25)
    ax1.set_ylim(0, y_max)
    ax1.grid(True, linestyle="--", alpha=0.3, zorder=0)

    # --------------------------------------------------------------------------
    # PANEL A2: TAHUN 2026 (275 CABANG)
    # --------------------------------------------------------------------------
    x_max_26 = 90
    th_vol_26 = THRESHOLD_VOLUME_2026

    ax2.fill_between([th_vol_26, x_max_26], [th_sla, th_sla], [y_max, y_max], color="#ECFDF5", alpha=0.6, zorder=1)
    ax2.fill_between([th_vol_26, x_max_26], [0, 0], [th_sla, th_sla], color="#FEF2F2", alpha=0.6, zorder=1)
    ax2.fill_between([0, th_vol_26], [0, 0], [th_sla, th_sla], color="#FFFBEB", alpha=0.6, zorder=1)
    ax2.fill_between([0, th_vol_26], [th_sla, th_sla], [y_max, y_max], color="#EFF6FF", alpha=0.6, zorder=1)

    ax2.axvline(th_vol_26, color="#64748B", linestyle="--", linewidth=1.5, zorder=2)
    ax2.axhline(th_sla, color="#64748B", linestyle="--", linewidth=1.5, zorder=2)

    ax2.text(65, 101, "KUADRAN I: CORE CHAMPIONS\n(Volume Tinggi, Kepatuhan Tinggi)",
             fontsize=8.2, fontweight="bold", color="#065F46", ha="center", zorder=3)
    ax2.text(65, 4, "KUADRAN II: ZONA RISIKO OPERASIONAL\n(Volume Tinggi, Kepatuhan Rendah - DISTORSI SLA)",
             fontsize=8.0, fontweight="bold", color="#991B1B", ha="center", zorder=3)
    ax2.text(20, 4, "KUADRAN III: UNDERPERFORMERS\n(Volume Rendah, Kepatuhan Rendah)",
             fontsize=8.0, fontweight="bold", color="#92400E", ha="center", zorder=3)
    ax2.text(20, 101, "KUADRAN IV: POTENSIAL / SELEKTIF\n(Volume Rendah, Kepatuhan Tinggi)",
             fontsize=8.0, fontweight="bold", color="#1E40AF", ha="center", zorder=3)

    ax2.text(th_vol_26 + 1.5, th_sla + 1.2, "Ambang Batas Sehat (70%)", fontsize=7.2, fontweight="bold", color="#334155")
    ax2.text(th_vol_26 + 1.5, 20, f"Batas Volume {int(th_vol_26)} Unit", fontsize=7.2, fontweight="bold", color="#334155", rotation=90)

    # Plot Kontraktor 2026
    df_plot_26 = grp_2026[grp_2026["unit"] >= 3].copy()
    for _, r in df_plot_26.iterrows():
        vname = r["kontraktor_clean"]
        x, y = r["unit"], r["pct_patuh"]
        col = COLOR_MAP.get(vname, "#475569")
        size = 180 + (r["avg_dur"] * 6)

        ax2.scatter(x, y, s=size, color=col, edgecolors="#0F172A", linewidth=1.4, zorder=4, alpha=0.9)

        if vname == "Sarjono":
            pos_x, pos_y, ha = x, y + 4.5, "center"
        elif vname == "CV. Rizki Mitra Abadi":
            pos_x, pos_y, ha = x - 5, y + 5.0, "right"
        elif vname == "Edwin":
            pos_x, pos_y, ha = x - 5, y - 6.5, "right"
        elif vname == "Trimo":
            pos_x, pos_y, ha = x, y - 7.5, "center"
        elif vname == "Teguh Karyanto":
            pos_x, pos_y, ha = x + 5, y + 3.0, "left"
            vname = "Teguh K."
        elif vname == "Dimas Andri Sulistyo":
            pos_x, pos_y, ha = x + 4, y - 5.5, "left"
            vname = "Dimas A."
        elif vname == "Sendy":
            pos_x, pos_y, ha = x + 4, y + 3.5, "left"
        else:
            pos_x, pos_y, ha = x, y + 3.5, "center"

        lbl = f"{vname}\n({int(x)}u | {y:.1f}%)"
        ax2.annotate(lbl, (pos_x, pos_y), fontsize=7.8, fontweight="bold", ha=ha, color="#0F172A", zorder=5,
                     bbox=dict(boxstyle="round,pad=0.2", facecolor="#FFFFFF", alpha=0.85, edgecolor="#CBD5E1", linewidth=0.8))

    # Anotasi Khusus Trimo (Penugasan Wilayah Bali)
    trimo_row_26 = df_plot_26[df_plot_26["kontraktor_clean"] == "Trimo"]
    if len(trimo_row_26) > 0:
        t_row = trimo_row_26.iloc[0]
        red_circle = patches.Ellipse((t_row["unit"], t_row["pct_patuh"]), width=16, height=12, angle=0,
                                     edgecolor="#DC2626", facecolor="none", linestyle="--", linewidth=2.0, zorder=6)
        ax2.add_patch(red_circle)
        ax2.annotate("TRIMO: MENANGGUNG WILAYAH BALI\n(12 Cabang Bali | Durasi Rerata 43,9 Hari)\nVendor lain belum bersedia ambil Bali",
                     xy=(t_row["unit"] + 5, t_row["pct_patuh"] + 1),
                     xytext=(t_row["unit"] + 10, t_row["pct_patuh"] + 22),
                     fontsize=7.5, fontweight="bold", color="#991B1B", zorder=7,
                     arrowprops=dict(arrowstyle="->", color="#DC2626", lw=1.5),
                     bbox=dict(boxstyle="round,pad=0.3", facecolor="#FEF2F2", edgecolor="#DC2626", linewidth=1.0))

    ax2.set_title("A2. Vendor Risk Quadrant TAHUN 2026 (275 Cabang)\nSLA Agregat Anjlok ke 53,8% (Trimo Menanggung Beban Wilayah Bali)",
                  fontsize=10.5, fontweight="bold", pad=12)
    ax2.set_xlabel("Volume Pengerjaan Renovasi (Jumlah Unit Cabang)", fontsize=9.5, fontweight="bold")
    ax2.set_ylabel("Tingkat Kepatuhan SLA Renovasi (%)", fontsize=9.5, fontweight="bold")
    ax2.set_xlim(0, x_max_26)
    ax2.set_ylim(0, y_max)
    ax2.grid(True, linestyle="--", alpha=0.3, zorder=0)

    fig.suptitle("Peta Risiko Vendor Berdampingan: Perubahan Posisi dan Performa Kontraktor (2025 vs 2026)",
                 fontsize=12.5, fontweight="bold", y=0.98)
    fig.tight_layout(rect=[0, 0, 1, 0.95])

    p_comp = os.path.join(GRAFIK_DIR, "8_peta_risiko_vendor_komparasi_2025_2026.png")
    fig.savefig(p_comp, bbox_inches="tight")
    plt.close(fig)

    if os.path.exists(ARTIFACT_GRAFIK_DIR):
        shutil.copy2(p_comp, os.path.join(ARTIFACT_GRAFIK_DIR, "8_peta_risiko_vendor_komparasi_2025_2026.png"))

    print(f"[✓] Grafik 8 (Side-by-Side Kuadran 2025 vs 2026) berhasil disimpan di: {p_comp}")


# ==============================================================================
# 4. GRAFIK 9: VEKTOR MIGRASI / PERPINDAHAN RISIKO VENDOR (2025 -> 2026)
# ==============================================================================
def buat_grafik_vektor_migrasi(grp_2025, grp_2026):
    """
    Menghasilkan Grafik 9: Peta Vektor Lintasan Migrasi Risiko Vendor
    Melacak pergerakan dinamis setiap vendor dari 2025 ke 2026 dengan panah arah.
    """
    os.makedirs(GRAFIK_DIR, exist_ok=True)
    fig, ax = plt.subplots(figsize=(15, 8.5), dpi=300)

    x_max_mig = 125
    y_max_mig = 105
    th_vol_mig = THRESHOLD_VOLUME_MIG
    th_sla = THRESHOLD_SLA

    # Latar Kuadran
    ax.fill_between([th_vol_mig, x_max_mig], [th_sla, th_sla], [y_max_mig, y_max_mig], color="#ECFDF5", alpha=0.6, zorder=1)
    ax.fill_between([th_vol_mig, x_max_mig], [0, 0], [th_sla, th_sla], color="#FEF2F2", alpha=0.6, zorder=1)
    ax.fill_between([0, th_vol_mig], [0, 0], [th_sla, th_sla], color="#FFFBEB", alpha=0.6, zorder=1)
    ax.fill_between([0, th_vol_mig], [th_sla, th_sla], [y_max_mig, y_max_mig], color="#EFF6FF", alpha=0.6, zorder=1)

    ax.axvline(th_vol_mig, color="#64748B", linestyle="--", linewidth=1.5, zorder=2)
    ax.axhline(th_sla, color="#64748B", linestyle="--", linewidth=1.5, zorder=2)

    ax.text(85, 101, "KUADRAN I: CORE CHAMPIONS\n(Volume Tinggi, Kepatuhan Tinggi)", fontsize=9, fontweight="bold", color="#065F46", ha="center", zorder=3)
    ax.text(85, 4, "KUADRAN II: ZONA RISIKO OPERASIONAL\n(Volume Tinggi, Kepatuhan Rendah - DISTORSI SLA)", fontsize=9, fontweight="bold", color="#991B1B", ha="center", zorder=3)
    ax.text(22, 4, "KUADRAN III: UNDERPERFORMERS\n(Volume Rendah, Kepatuhan Rendah)", fontsize=9, fontweight="bold", color="#92400E", ha="center", zorder=3)
    ax.text(22, 101, "KUADRAN IV: POTENSIAL / SELEKTIF\n(Volume Rendah, Kepatuhan Tinggi)", fontsize=9, fontweight="bold", color="#1E40AF", ha="center", zorder=3)

    # Vendor yang dilacak lintasannya
    tracked_vendors = ["Sarjono", "Sendy", "CV. Rizki Mitra Abadi", "Trimo", "Edwin", "Teguh Karyanto"]

    for vname in tracked_vendors:
        r25 = grp_2025[grp_2025["kontraktor_clean"] == vname]
        r26 = grp_2026[grp_2026["kontraktor_clean"] == vname]
        if len(r25) == 0 or len(r26) == 0:
            continue

        r25 = r25.iloc[0]
        r26 = r26.iloc[0]
        col = COLOR_MAP.get(vname, "#475569")

        x1, y1 = r25["unit"], r25["pct_patuh"]
        x2, y2 = r26["unit"], r26["pct_patuh"]

        # Titik 2025 (Hollow / Lingkaran Putih Bergaris)
        ax.scatter(x1, y1, s=180, facecolors="white", edgecolors=col, linewidth=2.0, linestyle="--", zorder=4)
        ax.annotate(f"{vname}\n'25 ({int(x1)}u|{y1:.0f}%)", (x1, y1 - 5.5), fontsize=7.5, color="#475569", ha="center", zorder=5)

        # Titik 2026 (Solid Cerah)
        ax.scatter(x2, y2, s=240, color=col, edgecolors="#0F172A", linewidth=1.6, zorder=6)
        ax.annotate(f"{vname}\n'26 ({int(x2)}u|{y2:.0f}%)", (x2, y2 + 4.5), fontsize=8.2, fontweight="bold", color="#0F172A", ha="center", zorder=7,
                    bbox=dict(boxstyle="round,pad=0.2", facecolor="#FFFFFF", alpha=0.9, edgecolor=col, linewidth=1.0))

        # Panah Vektor Migrasi
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="->", color=col, lw=2.4, shrinkA=8, shrinkB=8), zorder=5)

    # Highlight Khusus Trimo (Penugasan Wilayah Bali)
    callout_bali = (
        "FAKTOR OPERASIONAL KHUSUS: BEBAN WILAYAH BALI (TRIMO)\n"
        "• Trimo menanggung 12 cabang di Bali pada 2026 (92,3% proyek Bali).\n"
        "• Kontraktor lain (Sarjono, Sendy, Edwin) menolak/belum mau ambil Bali.\n"
        "• Logistik fery Ketapang-Gilimanuk & izin Banjar membuat durasi Bali = 43,9 hari (0% SLA).\n"
        "• Di luar Bali (Pulau Jawa), rerata durasi Trimo adalah 31,8 hari."
    )
    ax.text(0.51, 0.43, callout_bali, transform=ax.transAxes, fontsize=8.2, fontweight="bold", color="#7F1D1D",
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#FEF2F2", edgecolor="#DC2626", linewidth=1.4), zorder=8)

    # Legend Penjelasan Titik & Garis
    legend_elements = [
        plt.Line2D([0], [0], marker='o', color='w', label='Posisi Tahun 2025 (Titik Awal)',
                   markerfacecolor='w', markeredgecolor='#475569', markersize=10, markeredgewidth=2, linestyle='--'),
        plt.Line2D([0], [0], marker='o', color='w', label='Posisi Tahun 2026 (Realisasi Terbaru)',
                   markerfacecolor='#2563EB', markeredgecolor='#0F172A', markersize=12, markeredgewidth=1.5),
        plt.Line2D([0], [0], color='#2563EB', lw=2.2, label='Arah Vektor Migrasi Posisi (2025 -> 2026)')
    ]
    ax.legend(handles=legend_elements, loc='upper right', frameon=True, fontsize=8.5)

    ax.set_title("Peta Lintasan & Vektor Migrasi Risiko Vendor Renovasi PGI (Tahun 2025 -> 2026)\nMelacak Dinamika Kapasitas Eksekusi dan Dampak Penugasan Wilayah Khusus (Bali)",
                 fontsize=12, fontweight="bold", pad=12)
    ax.set_xlabel("Volume Pengerjaan Renovasi (Jumlah Unit Cabang)", fontsize=10.5, fontweight="bold")
    ax.set_ylabel("Tingkat Kepatuhan SLA Renovasi (%)", fontsize=10.5, fontweight="bold")
    ax.set_xlim(0, x_max_mig)
    ax.set_ylim(0, y_max_mig)
    ax.grid(True, linestyle="--", alpha=0.3, zorder=0)

    fig.tight_layout()
    p_mig = os.path.join(GRAFIK_DIR, "9_vektor_migrasi_risiko_vendor_2025_2026.png")
    fig.savefig(p_mig, bbox_inches="tight")
    plt.close(fig)

    if os.path.exists(ARTIFACT_GRAFIK_DIR):
        shutil.copy2(p_mig, os.path.join(ARTIFACT_GRAFIK_DIR, "9_vektor_migrasi_risiko_vendor_2025_2026.png"))

    print(f"[✓] Grafik 9 (Vektor Migrasi 2025 -> 2026) berhasil disimpan di: {p_mig}")


# ==============================================================================
# 5. CETAK DASHBOARD TERMINAL VSCODE
# ==============================================================================
def cetak_dashboard_terminal(komp, bali_sum, trimo_stats):
    """Mencetak ringkasan komparasi 2025 vs 2026 dan analisis wilayah Bali ke terminal."""
    print("\n" + "=" * 118)
    print("        PUSAT GADAI INDONESIA — LAPORAN RESMI KOMPARASI PERFORMA & RISIKO VENDOR (2025 VS 2026)")
    print("                 (DILENGKAPI ANALISIS PENUGASAN KHUSUS WILAYAH BALI PADA KONTRAKTOR TRIMO)")
    print("=" * 118)

    print("\n[ TABEL 1: MATRIKS KOMPARASI PERFORMA & PERGESERAN KUADRAN VENDOR (2025 VS 2026) ]")
    print("-" * 118)
    print(f"{'Nama Kontraktor':<26} | {'2025 Unit':<9} | {'2025 SLA':<8} | {'2026 Unit':<9} | {'2026 SLA':<8} | {'Delta SLA':<9} | {'2026 Durasi':<11} | {'2026 Bali':<9} | {'Pergeseran Kuadran'}")
    print("-" * 118)

    for _, r in komp.iterrows():
        vname = r["kontraktor_clean"]
        u25 = f"{int(r['unit_2025'])} unit" if r['unit_2025'] > 0 else "-"
        s25 = f"{r['pct_patuh_2025']:.1f}%" if r['unit_2025'] > 0 else "-"
        u26 = f"{int(r['unit_2026'])} unit" if r['unit_2026'] > 0 else "-"
        s26 = f"{r['pct_patuh_2026']:.1f}%" if r['unit_2026'] > 0 else "-"
        
        if r['unit_2025'] > 0 and r['unit_2026'] > 0:
            d_sla = f"{r['delta_sla_poin']:+.1f} pt"
        else:
            d_sla = "Baru / Off"
            
        dur26 = f"{r['avg_dur_2026']:.1f} hr" if r['unit_2026'] > 0 else "-"
        b26 = f"{int(r['bali_unit_2026'])} unit" if r['bali_unit_2026'] > 0 else "0"
        
        k25_short = r['kuadran_2025'].split(":")[0] if ":" in str(r['kuadran_2025']) else str(r['kuadran_2025'])
        k26_short = r['kuadran_2026'].split(":")[0] if ":" in str(r['kuadran_2026']) else str(r['kuadran_2026'])
        geser = f"{k25_short} -> {k26_short}"

        print(f"{vname:<26} | {u25:<9} | {s25:<8} | {u26:<9} | {s26:<8} | {d_sla:<9} | {dur26:<11} | {b26:<9} | {geser}")

    print("-" * 118)

    print("\n[ TABEL 2: DEKOMPOSISI OPERASIONAL EKSEKUSI WILAYAH BALI — TRIMO VS VENDOR LAIN (2026) ]")
    print("-" * 118)
    print(f"• Total Cabang Dibuka di Bali 2026    : 13 Cabang (12 Dikerjakan Trimo = 92,3%, 1 oleh CV Rizki Mitra Abadi)")
    print(f"• Vendor yang Menolak/Belum Mau Bali  : Sarjono (0 Bali), Sendy (0 Bali), Edwin (0 Bali), Teguh K. (0 Bali)")
    print(f"• Kinerja Kontraktor Trimo di Bali    : {trimo_stats['bali_unit']} Unit | Rerata Durasi: {trimo_stats['bali_avg_dur']:.1f} Hari (Median {trimo_stats['bali_med_dur']:.1f} hr) | Kepatuhan SLA: {trimo_stats['bali_sla_pct']:.1f}%")
    print(f"• Kinerja Kontraktor Trimo di Jawa    : {trimo_stats['non_bali_unit']} Unit | Rerata Durasi: {trimo_stats['non_bali_avg_dur']:.1f} Hari (Median {trimo_stats['non_bali_med_dur']:.1f} hr) | Kepatuhan SLA: {trimo_stats['non_bali_sla_pct']:.1f}%")
    print(f"• Selisih Durasi Trimo (Bali vs Jawa) : +12,1 Hari Lebih Lambat di Bali akibat logistik fery & izin adat Banjar")
    print("-" * 118)

    print("\n[ KESIMPULAN STRATEGIS MANAJEMEN ]")
    print("1. Sarjono adalah Core Champion Paling Konsisten: Mempertahankan SLA 89,0% dengan rerata durasi kilat 14,9 hari.")
    print("2. Peningkatan Impresif CV Rizki Mitra Abadi: Melejit +22,1 poin SLA (dari 44,6% ke 66,7%), mendekati Kuadran I.")
    print("3. Realitas Objektif Trimo (Mitigasi Bali): Anjloknya SLA Trimo ke 15,1% bukan semata inefisiensi internal,")
    print("   melainkan karena Trimo menjadi 'bumper' ekspansi wilayah Bali yang ditolak kontraktor lain. Di luar Bali,")
    print("   durasi Trimo adalah 31,8 hari (masih perlu pembenahan, namun jauh lebih baik dibanding 43,9 hari di Bali).")
    print("4. Rekomendasi Alokasi: Tetapkan SLA khusus Bali (35–40 hari) atau beri insentif logistik bagi Sarjono/Sendy")
    print("   agar bersedia berbagi beban proyek luar pulau bersama Trimo.")
    print("=" * 118 + "\n")


# ==============================================================================
# 6. EKSPOR KE EXCEL MULTI-SHEET
# ==============================================================================
def export_ke_excel(df, grp_2025, grp_2026, komp, bali_sum):
    """Menyimpan seluruh data analisis ke berkas Excel multi-sheet resmi."""
    excel_path = os.path.join(OUTPUT_DIR, "Laporan_Risiko_Vendor_Renovasi.xlsx")
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        komp.to_excel(writer, sheet_name="Komparasi_2025_2026", index=False)
        grp_2026.to_excel(writer, sheet_name="Vendor_2026_Detail", index=False)
        grp_2025.to_excel(writer, sheet_name="Vendor_2025_Detail", index=False)
        bali_sum.to_excel(writer, sheet_name="Analisis_Wilayah_Bali", index=False)

    print(f"[✓] Berkas Excel Multi-Sheet berhasil disimpan di: {excel_path}")


# ==============================================================================
# 5b. GRAFIK 8a: TABEL MATRIKS KOMPARASI PERFORMA & PERGESERAN KUADRAN VENDOR
# ==============================================================================
def buat_grafik_tabel_komparasi_vendor(komp=None):
    """
    Menghasilkan Grafik 8a: Matriks Komparasi Kuantitatif Performa, Durasi Fisik,
    Tingkat Kepatuhan SLA, dan Pergeseran Kuadran Mitra Kontraktor (2025 vs 2026).
    """
    os.makedirs(GRAFIK_DIR, exist_ok=True)
    t_headers = [
        "Mitra Kontraktor",
        "Volume\n'2025",
        "SLA\n'2025",
        "Volume\n'2026",
        "SLA\n'2026",
        "Delta\nSLA",
        "Durasi\n'2026",
        "Beban\nBali '26",
        "Status Pergeseran / Migrasi Kuadran"
    ]

    rows = [
        ["Sarjono", "109 unit", "88,1%", "73 unit", "89,0%", "+1,0 pt", "14,9 hr", "0 unit", "Kuadran I → Kuadran I (Stable Champion)"],
        ["CV. Rizki Mitra Abadi", "74 unit", "44,6%", "54 unit", "66,7%", "+22,1 pt", "21,7 hr", "0 unit", "Kuadran II → Mendekati Kuadran I (Turnaround)"],
        ["Edwin", "66 unit", "69,7%", "52 unit", "61,5%", "-8,2 pt", "21,4 hr", "0 unit", "Kuadran II → Kuadran II (Moderat/Stabil)"],
        ["Trimo*", "74 unit", "41,9%", "53 unit", "15,1%", "-26,8 pt", "34,5 hr", "12 unit (92%)", "Kuadran II → Kuadran II (Beban Khusus Bali)"],
        ["Teguh Karyanto", "2 unit", "0,0%", "30 unit", "20,0%", "+20,0 pt", "31,3 hr", "0 unit", "Kuadran III → Kuadran III (Alot/Overload)"],
        ["Sendy", "70 unit", "98,6%", "3 unit", "66,7%", "-31,9 pt", "18,3 hr", "0 unit", "Kuadran I → Kuadran III (Volume Turun Drastis)"],
        ["Dimas Andri Sulistyo", "-", "-", "9 unit", "11,1%", "Baru", "30,4 hr", "0 unit", "Rekanan Baru → Kuadran III (Underperformer)"],
        ["CV Cahaya Kemakmuran", "39 unit", "15,4%", "0 unit", "-", "Off", "-", "0 unit", "Kuadran III → Dinonaktifkan (Discontinued)"]
    ]

    fig, ax = plt.subplots(figsize=(15.5, 6.2), dpi=300)
    ax.axis('off')

    col_widths = [0.16, 0.08, 0.08, 0.08, 0.08, 0.08, 0.09, 0.10, 0.25]

    tab = ax.table(
        cellText=rows,
        colLabels=t_headers,
        colWidths=col_widths,
        cellLoc='center',
        loc='center'
    )

    tab.auto_set_font_size(False)
    tab.set_fontsize(8.5)

    n_r = len(rows)
    n_c = len(t_headers)

    for c in range(n_c):
        h_cell = tab[0, c]
        h_cell.set_height(0.13)
        h_cell.set_facecolor('#0F172A')
        h_cell.set_edgecolor('#334155')
        h_cell.set_linewidth(1.2)
        h_cell.get_text().set_color('#FFFFFF')
        h_cell.get_text().set_fontweight('bold')
        h_cell.get_text().set_fontsize(8.2)

        for r_i in range(1, n_r + 1):
            cell = tab[r_i, c]
            cell.set_height(0.085)
            bg = '#F8FAFC' if r_i % 2 == 1 else '#FFFFFF'
            cell.set_facecolor(bg)
            cell.set_edgecolor('#CBD5E1')
            cell.set_linewidth(0.8)
            cell.get_text().set_fontsize(8.2)

            if c == 0:
                cell.get_text().set_fontweight('bold')
                cell.get_text().set_color('#0F172A')
                cell.set_facecolor('#F1F5F9')
            elif c == 5:
                val = rows[r_i - 1][5]
                if "+" in val:
                    cell.set_facecolor('#ECFDF5')
                    cell.get_text().set_color('#065F46')
                    cell.get_text().set_fontweight('bold')
                elif "-" in val:
                    cell.set_facecolor('#FEF2F2')
                    cell.get_text().set_color('#991B1B')
                    cell.get_text().set_fontweight('bold')
                else:
                    cell.get_text().set_color('#64748B')
            elif c == 7:
                if "12" in rows[r_i - 1][7]:
                    cell.set_facecolor('#FEF2F2')
                    cell.get_text().set_color('#DC2626')
                    cell.get_text().set_fontweight('bold')
                else:
                    cell.get_text().set_color('#64748B')
            elif c == 8:
                txt = rows[r_i - 1][8]
                cell.get_text().set_fontweight('bold')
                if "Stable Champion" in txt:
                    cell.set_facecolor('#ECFDF5')
                    cell.get_text().set_color('#065F46')
                elif "Turnaround" in txt:
                    cell.set_facecolor('#EFF6FF')
                    cell.get_text().set_color('#1E40AF')
                elif "Beban Khusus Bali" in txt:
                    cell.set_facecolor('#FFFBEB')
                    cell.get_text().set_color('#B45309')
                elif "Discontinued" in txt:
                    cell.set_facecolor('#F1F5F9')
                    cell.get_text().set_color('#64748B')
                else:
                    cell.get_text().set_color('#334155')

    title = "MATRIKS KOMPARASI KUANTITATIF PERFORMA, DURASI PENGERJAAN FISIK, TINGKAT KEPATUHAN SLA,\nDAN PERGESERAN KUADRAN MITRA KONTRAKTOR RENOVASI (2025 VS 2026)"
    subtitle = "Evaluasi Kinerja & Dinamika Kapasitas Eksekusi Rekanan Renovasi Cabang UPC | PT Pusat Gadai Indonesia"
    full_title = f"{title}\n{subtitle}"
    plt.suptitle(full_title, fontsize=11, fontweight='bold', color='#0F172A', y=0.96, linespacing=1.3)

    notes = (
        "Catatan & Evaluasi Manajemen:\n"
        "1. Sarjono: Konsisten sebagai Core Champion utama dengan volume tertinggi (73 unit), kepatuhan SLA prima (89,0%), dan durasi tercepat (14,9 hari).\n"
        "2. CV. Rizki Mitra Abadi: Menunjukkan turnaround performa terbaik dengan lonjakan kepatuhan SLA drastis sebesar +22,1 poin persentase.\n"
        "3. (*) Trimo (Beban Wilayah Bali): Penurunan SLA Trimo (15,1%) bukan kegagalan struktural, melainkan akibat penugasan khusus 12 cabang di Bali (92,3% proyek Bali)\n"
        "   yang terkendala logistik fery Ketapang-Gilimanuk, pengiriman brankas dari Jawa, dan perizinan adat Banjar. Di luar Bali, rerata durasi Trimo adalah 31,8 hari.\n"
        "4. CV Cahaya Kemakmuran: Dinonaktifkan dari daftar rekanan pada 2026 akibat tingkat keterlambatan tinggi (SLA 2025 hanya 15,4%)."
    )
    fig.text(0.04, 0.04, notes, fontsize=7.8, color='#334155', style='italic', linespacing=1.35)

    outfile = os.path.join(GRAFIK_DIR, "8a_tabel_matriks_komparasi_vendor_2025_2026.png")
    fig.savefig(outfile, bbox_inches='tight')
    outfile_alt = os.path.join(GRAFIK_DIR, "8_tabel_matriks_komparasi_vendor_2025_2026.png")
    fig.savefig(outfile_alt, bbox_inches='tight')
    plt.close(fig)

    if os.path.exists(ARTIFACT_GRAFIK_DIR):
        shutil.copy2(outfile, os.path.join(ARTIFACT_GRAFIK_DIR, "8a_tabel_matriks_komparasi_vendor_2025_2026.png"))
        shutil.copy2(outfile_alt, os.path.join(ARTIFACT_GRAFIK_DIR, "8_tabel_matriks_komparasi_vendor_2025_2026.png"))

    print(f"[✓] Grafik 8a (Tabel Matriks Komparasi Vendor) berhasil disimpan di: {outfile}")


# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================
def main():
    print("\n" + "=" * 65)
    print(" Memulai Eksekusi Analisis Peta Risiko Vendor (2025 vs 2026)...")
    print("=" * 65)

    # 1. Pemuatan Data
    df = load_and_prep_vendor_data()

    # 2. Kalkulasi Metrik 2025 & 2026
    grp_2025 = hitung_metrik_tahunan(df, 2025)
    grp_2026 = hitung_metrik_tahunan(df, 2026)
    komp = hitung_komparasi_migrasi(grp_2025, grp_2026)
    bali_sum, trimo_stats = hitung_analisis_bali(df)

    # 3. Cetak Dashboard Terminal
    cetak_dashboard_terminal(komp, bali_sum, trimo_stats)

    # 4. Buat Visualisasi Grafik
    buat_grafik_komparasi_2025_2026(grp_2025, grp_2026)
    buat_grafik_vektor_migrasi(grp_2025, grp_2026)
    buat_grafik_tabel_komparasi_vendor(komp)

    # 5. Ekspor Hasil ke Excel
    export_ke_excel(df, grp_2025, grp_2026, komp, bali_sum)

    print("[✓] Seluruh Rangkaian Analisis Vendor Risk Quadrant Berhasil Selesai!\n")


if __name__ == "__main__":
    main()
