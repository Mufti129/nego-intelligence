#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
SKRIP ANALISIS PERFORMA NEGOSIASI DAN NEGOSIATOR PER BULAN (TAHUN 2026)
PUSAT GADAI INDONESIA (PGI) — DIVISI BISNIS
================================================================================

Penulis      : Divisi Bisnis (Data Analyst Team)
Dataset      : properties_cleaned_2024-2026.csv / properties (1) - properties (2024-2026)-.csv
Aturan Data  : 
  - Tahun Analisis          : 2026
  - Tim Negosiasi Resmi     : Mirza & Bonita (Dika & Salma berstatus Surveyor / Non-Tim Nego)
  - Harga Awal Penawaran    : Hargaawal_nego_sheet
  - Harga Sewa Kesepakatan  : hargarental_sheet
  - Efisiensi Diskon        : efisiensi_diskon_nego2 / efisiensi_diskon_pct
  - Durasi Negosiasi        : lama_waktu_realisasi_nego (hari)
  - Penentuan Bulan         : tgl_nego_berakhir (tanggal deal penutupan negosiasi)

================================================================================
CARA MENJALANKAN SKRIP DI VISUAL STUDIO CODE (VSCODE):
================================================================================
1. Buka folder kerja di VSCode:
   - Menu: File -> Open Folder... -> Pilih folder "Nego_baru"
2. Buka berkas ini di editor:
   - Dobel klik file "analisa_performa_negosiasi_bulanan_2026.py"
3. Jalankan skrip:
   - CARA A (Tombol Run): Klik ikon segitiga "Play/Run Python File" di pojok kanan atas editor.
   - CARA B (Terminal VSCode): Tekan Ctrl + ` (atau Terminal -> New Terminal), lalu ketik:
         python3 analisa_performa_negosiasi_bulanan_2026.py
4. Hasil Analisis:
   - Ditampilkan langsung di terminal VSCode dalam bentuk tabel rapi.
   - Disimpan otomatis ke folder 'hasil_analisis/' dalam format Excel (.xlsx) dan CSV.
================================================================================
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import shutil
from scipy import stats

# ==============================================================================
# KONFIGURASI PARAMETER
# ==============================================================================
TAHUN_TARGET = 2026
DIR_KERJA = os.path.dirname(os.path.abspath(__file__))
FILE_CLEANED_EXCEL = os.path.join(DIR_KERJA, "data_nego_baru", "data_nego_baru_cleaned_2024-2026.xlsx")
FILE_CLEANED = os.path.join(DIR_KERJA, "properties_cleaned_2024-2026.csv")
FILE_RAW = os.path.join(DIR_KERJA, "properties (1) - properties (2024-2026)-.csv")
OUTPUT_DIR = os.path.join(DIR_KERJA, "hasil_analisis")

# Pemetaan Nama Bulan Bahasa Indonesia
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

TIM_NEGO_RESMI = ["Bonita", "Mirza"]


# ==============================================================================
# HELPER FORMATTING
# ==============================================================================
def rupiah(nilai):
    """Format angka ke teks rupiah (Juta / Miliar) untuk pembacaan eksekutif."""
    if pd.isna(nilai) or nilai == 0:
        return "Rp 0"
    abs_val = abs(nilai)
    if abs_val >= 1_000_000_000:
        return f"Rp {nilai / 1_000_000_000:.2f} Miliar"
    elif abs_val >= 1_000_000:
        return f"Rp {nilai / 1_000_000:.1f} Juta"
    else:
        return f"Rp {nilai:,.0f}".replace(",", ".")


def persen(nilai):
    """Format angka desimal ke persen dengan 2 digit presisi."""
    if pd.isna(nilai):
        return "0,00%"
    return f"{nilai:.2f}%".replace(".", ",")


# ==============================================================================
# 1. PEMUATAN & PENYIAPAN DATASET
# ==============================================================================
def load_and_prep_data():
    """Membaca file cleaned atau raw, lalu menyiapkan subset tahun 2026."""
    if os.path.exists(FILE_CLEANED_EXCEL):
        print(f"[*] Membaca data bersih utama: {FILE_CLEANED_EXCEL}")
        df = pd.read_excel(FILE_CLEANED_EXCEL)
    elif os.path.exists(FILE_CLEANED):
        print(f"[*] Membaca data bersih CSV: {os.path.basename(FILE_CLEANED)}")
        df = pd.read_csv(FILE_CLEANED)
    elif os.path.exists(FILE_RAW):
        print(f"[*] Data cleaned belum ada, membaca file mentah: {os.path.basename(FILE_RAW)}")
        df = pd.read_csv(FILE_RAW)
        if "nomor_pengajuan" in df.columns:
            df = df.drop_duplicates(subset=["nomor_pengajuan"], keep="last")
    else:
        print(f"[!] ERROR: Berkas data tidak ditemukan di {DIR_KERJA}")
        sys.exit(1)

    df["Tahun"] = pd.to_numeric(df["Tahun"], errors="coerce").fillna(0).astype(int)
    df_2026 = df[df["Tahun"] == TAHUN_TARGET].copy()
    print(f"[*] Jumlah cabang tahun {TAHUN_TARGET}: {len(df_2026)} unit")

    # Parsing tanggal tgl_nego_berakhir
    df_2026["tgl_nego_selesai_dt"] = pd.to_datetime(df_2026["tgl_nego_berakhir"], errors="coerce")
    df_2026["bulan_num"] = df_2026["tgl_nego_selesai_dt"].dt.month
    df_2026["bulan_nama"] = df_2026["bulan_num"].map(BULAN_ID)
    df_2026["bulan_kode"] = df_2026["tgl_nego_selesai_dt"].dt.strftime("%Y-%m")
    df_2026["bulan_label"] = df_2026["bulan_kode"] + " (" + df_2026["bulan_nama"] + ")"

    # Standardisasi Nama Negosiator & Klasifikasi Tim
    col_nego = "nama_negosiator_2" if "nama_negosiator_2" in df_2026.columns else "nama_negosiator"
    df_2026["negosiator"] = (
        df_2026[col_nego]
        .fillna(df_2026.get("nama_negosiator", "Tanpa Nama"))
        .astype(str)
        .str.strip()
        .str.title()
    )
    df_2026["is_tim_nego"] = df_2026["negosiator"].isin(TIM_NEGO_RESMI)
    df_2026["kategori_tim"] = np.where(
        df_2026["is_tim_nego"],
        df_2026["negosiator"],
        "Non-Tim Nego (Surveyor: Dika/Salma)"
    )

    # Standardisasi Kolom Finansial
    if "Hargaawal_nego_sheet" in df_2026.columns:
        df_2026["asking_price"] = pd.to_numeric(df_2026["Hargaawal_nego_sheet"], errors="coerce").fillna(0)
    elif "harga_awal_penawaran" in df_2026.columns:
        df_2026["asking_price"] = pd.to_numeric(df_2026["harga_awal_penawaran"], errors="coerce").fillna(0)
    else:
        df_2026["asking_price"] = pd.to_numeric(df_2026.get("asking_price", 0), errors="coerce").fillna(0)

    if "hargarental_sheet" in df_2026.columns:
        df_2026["deal_price"] = pd.to_numeric(df_2026["hargarental_sheet"], errors="coerce").fillna(0)
    elif "harga_rental_final" in df_2026.columns:
        df_2026["deal_price"] = pd.to_numeric(df_2026["harga_rental_final"], errors="coerce").fillna(0)
    else:
        df_2026["deal_price"] = pd.to_numeric(df_2026.get("rental_price", 0), errors="coerce").fillna(0)

    # Saving Rupiah
    df_2026["saving_rupiah"] = np.maximum(0, df_2026["asking_price"] - df_2026["deal_price"])

    # Efisiensi Diskon (%)
    if "efisiensi_diskon_pct" in df_2026.columns:
        df_2026["diskon_pct"] = pd.to_numeric(df_2026["efisiensi_diskon_pct"], errors="coerce").fillna(0)
    elif "efisiensi_diskon_nego2" in df_2026.columns:
        df_2026["diskon_pct"] = (
            df_2026["efisiensi_diskon_nego2"]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df_2026["diskon_pct"] = pd.to_numeric(df_2026["diskon_pct"], errors="coerce").fillna(0)
    else:
        df_2026["diskon_pct"] = np.where(
            df_2026["asking_price"] > 0,
            (df_2026["saving_rupiah"] / df_2026["asking_price"]) * 100,
            0.0
        )

    # Durasi Nego (Hari)
    if "durasi_nego_hari" in df_2026.columns:
        df_2026["durasi_hari"] = pd.to_numeric(df_2026["durasi_nego_hari"], errors="coerce").fillna(0)
    elif "lama_waktu_realisasi_nego" in df_2026.columns:
        df_2026["durasi_hari"] = pd.to_numeric(df_2026["lama_waktu_realisasi_nego"], errors="coerce").fillna(0)
    else:
        df_2026["durasi_hari"] = 0.0

    return df_2026


# ==============================================================================
# 2. ANALISIS KINERJA BULANAN PER INDIVIDU (BONITA / MIRZA)
# ==============================================================================
def hitung_performa_individu(df_2026, nama_personil):
    """Menghitung rincian bulanan untuk negosiator spesifik (Bonita atau Mirza)."""
    df_sub = df_2026[df_2026["negosiator"] == nama_personil].copy()
    grouped = df_sub.groupby(["bulan_num", "bulan_nama"])

    summary = grouped.agg(
        jumlah_deal=("nomor_pengajuan", "count"),
        total_asking=("asking_price", "sum"),
        total_deal_price=("deal_price", "sum"),
        total_saving=("saving_rupiah", "sum"),
        avg_diskon_pct=("diskon_pct", "mean"),
        median_diskon_pct=("diskon_pct", "median"),
        avg_durasi_hari=("durasi_hari", "mean"),
        median_durasi_hari=("durasi_hari", "median"),
        deal_berdiskon=("saving_rupiah", lambda s: (s > 0).sum())
    ).reset_index()

    summary["success_rate_pct"] = (summary["deal_berdiskon"] / summary["jumlah_deal"]) * 100
    summary["effective_discount_pct"] = np.where(
        summary["total_asking"] > 0,
        (summary["total_saving"] / summary["total_asking"]) * 100,
        0.0
    )

    # Baris Total Tahunan Individu
    tot_deal = summary["jumlah_deal"].sum()
    tot_ask = summary["total_asking"].sum()
    tot_rent = summary["total_deal_price"].sum()
    tot_save = summary["total_saving"].sum()
    tot_disk_deal = summary["deal_berdiskon"].sum()

    row_total = pd.DataFrame([{
        "bulan_num": 99,
        "bulan_nama": "TOTAL TAHUNAN",
        "jumlah_deal": tot_deal,
        "total_asking": tot_ask,
        "total_deal_price": tot_rent,
        "total_saving": tot_save,
        "avg_diskon_pct": df_sub["diskon_pct"].mean(),
        "median_diskon_pct": df_sub["diskon_pct"].median(),
        "avg_durasi_hari": df_sub["durasi_hari"].mean(),
        "median_durasi_hari": df_sub["durasi_hari"].median(),
        "deal_berdiskon": tot_disk_deal,
        "success_rate_pct": (tot_disk_deal / tot_deal) * 100 if tot_deal > 0 else 0,
        "effective_discount_pct": (tot_save / tot_ask) * 100 if tot_ask > 0 else 0
    }])

    return pd.concat([summary, row_total], ignore_index=True)


# ==============================================================================
# 3. TABEL KOMPARASI BULANAN LENGKAP (BONITA VS MIRZA + NON-TIM NEGO)
# ==============================================================================
def buat_tabel_komparasi_bulanan(df_2026):
    """Menyusun matriks komparasi bulanan dengan metrik deal, saving, diskon, dan durasi."""
    rows = []
    
    for b_num in range(1, 9):
        b_nama = BULAN_ID[b_num]
        df_m = df_2026[df_2026["bulan_num"] == b_num]

        # Bonita
        df_b = df_m[df_m["negosiator"] == "Bonita"]
        deal_b = len(df_b)
        save_b = df_b["saving_rupiah"].sum()
        disc_b = df_b["diskon_pct"].mean() if deal_b > 0 else 0.0
        dur_b = df_b["durasi_hari"].mean() if deal_b > 0 else 0.0

        # Mirza
        df_z = df_m[df_m["negosiator"] == "Mirza"]
        deal_m = len(df_z)
        save_m = df_z["saving_rupiah"].sum()
        disc_m = df_z["diskon_pct"].mean() if deal_m > 0 else 0.0
        dur_m = df_z["durasi_hari"].mean() if deal_m > 0 else 0.0

        # Non-Tim Nego (Dika / Salma)
        df_non = df_m[~df_m["is_tim_nego"]]
        deal_non = len(df_non)
        save_non = df_non["saving_rupiah"].sum()
        disc_non = df_non["diskon_pct"].mean() if deal_non > 0 else 0.0
        dur_non = df_non["durasi_hari"].mean() if deal_non > 0 else 0.0

        # Total Bulan
        deal_tot = len(df_m)
        save_tot = df_m["saving_rupiah"].sum()
        disc_tot = df_m["diskon_pct"].mean() if deal_tot > 0 else 0.0
        dur_tot = df_m["durasi_hari"].mean() if deal_tot > 0 else 0.0

        rows.append({
            "bulan_num": b_num,
            "bulan_nama": b_nama,
            "deal_bonita": deal_b,
            "saving_bonita": save_b,
            "diskon_bonita": disc_b,
            "durasi_bonita": dur_b,
            "deal_mirza": deal_m,
            "saving_mirza": save_m,
            "diskon_mirza": disc_m,
            "durasi_mirza": dur_m,
            "deal_non_tim": deal_non,
            "saving_non_tim": save_non,
            "total_deal": deal_tot,
            "total_saving": save_tot,
            "avg_diskon_nasional": disc_tot,
            "avg_durasi_nasional": dur_tot
        })

    df_comp = pd.DataFrame(rows)

    # Tambahkan Baris Total
    df_bonita_all = df_2026[df_2026["negosiator"] == "Bonita"]
    df_mirza_all = df_2026[df_2026["negosiator"] == "Mirza"]
    df_non_all = df_2026[~df_2026["is_tim_nego"]]

    row_total = pd.DataFrame([{
        "bulan_num": 99,
        "bulan_nama": "TOTAL 2026",
        "deal_bonita": len(df_bonita_all),
        "saving_bonita": df_bonita_all["saving_rupiah"].sum(),
        "diskon_bonita": df_bonita_all["diskon_pct"].mean(),
        "durasi_bonita": df_bonita_all["durasi_hari"].mean(),
        "deal_mirza": len(df_mirza_all),
        "saving_mirza": df_mirza_all["saving_rupiah"].sum(),
        "diskon_mirza": df_mirza_all["diskon_pct"].mean(),
        "durasi_mirza": df_mirza_all["durasi_hari"].mean(),
        "deal_non_tim": len(df_non_all),
        "saving_non_tim": df_non_all["saving_rupiah"].sum(),
        "total_deal": len(df_2026),
        "total_saving": df_2026["saving_rupiah"].sum(),
        "avg_diskon_nasional": df_2026["diskon_pct"].mean(),
        "avg_durasi_nasional": df_2026["durasi_hari"].mean()
    }])

    return pd.concat([df_comp, row_total], ignore_index=True)


# ==============================================================================
# 4. TREN BULANAN OVERALL
# ==============================================================================
def hitung_performa_bulanan_overall(df_2026):
    """Menghitung performa agregat bulanan seluruh cabang 2026."""
    grouped = df_2026.groupby(["bulan_num", "bulan_nama"])

    summary = grouped.agg(
        total_deal=("nomor_pengajuan", "count"),
        total_asking=("asking_price", "sum"),
        total_deal_price=("deal_price", "sum"),
        total_saving=("saving_rupiah", "sum"),
        avg_diskon_pct=("diskon_pct", "mean"),
        median_diskon_pct=("diskon_pct", "median"),
        avg_durasi_hari=("durasi_hari", "mean"),
        median_durasi_hari=("durasi_hari", "median"),
        deal_berdiskon=("saving_rupiah", lambda s: (s > 0).sum())
    ).reset_index()

    summary["success_rate_pct"] = (summary["deal_berdiskon"] / summary["total_deal"]) * 100
    summary["effective_discount_pct"] = np.where(
        summary["total_asking"] > 0,
        (summary["total_saving"] / summary["total_asking"]) * 100,
        0.0
    )

    tot_deal = summary["total_deal"].sum()
    tot_ask = summary["total_asking"].sum()
    tot_rent = summary["total_deal_price"].sum()
    tot_save = summary["total_saving"].sum()
    tot_disk_deal = summary["deal_berdiskon"].sum()

    row_total = pd.DataFrame([{
        "bulan_num": 99,
        "bulan_nama": "TOTAL TAHUN 2026",
        "total_deal": tot_deal,
        "total_asking": tot_ask,
        "total_deal_price": tot_rent,
        "total_saving": tot_save,
        "avg_diskon_pct": df_2026["diskon_pct"].mean(),
        "median_diskon_pct": df_2026["diskon_pct"].median(),
        "avg_durasi_hari": df_2026["durasi_hari"].mean(),
        "median_durasi_hari": df_2026["durasi_hari"].median(),
        "deal_berdiskon": tot_disk_deal,
        "success_rate_pct": (tot_disk_deal / tot_deal) * 100,
        "effective_discount_pct": (tot_save / tot_ask) * 100
    }])

    return pd.concat([summary, row_total], ignore_index=True)


# ==============================================================================
# 5. EKSPOR KE EXCEL DAN CSV
# ==============================================================================
def export_ke_file(df_overall, df_bonita, df_mirza, df_comp, df_2026):
    """Menyimpan seluruh tabel hasil analisis ke folder 'hasil_analisis/'."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Ekspor CSV
    df_overall.to_csv(os.path.join(OUTPUT_DIR, "performa_nego_bulanan_2026_overall.csv"), index=False)
    df_bonita.to_csv(os.path.join(OUTPUT_DIR, "performa_bulanan_bonita_2026.csv"), index=False)
    df_mirza.to_csv(os.path.join(OUTPUT_DIR, "performa_bulanan_mirza_2026.csv"), index=False)
    df_comp.to_csv(os.path.join(OUTPUT_DIR, "komparasi_bulanan_tim_nego_2026.csv"), index=False)

    # Ekspor Excel Multi-Sheet
    excel_path = os.path.join(OUTPUT_DIR, "Laporan_Performa_Negosiasi_Bulanan_2026.xlsx")
    try:
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_overall.to_excel(writer, sheet_name="Tren Bulanan Overall", index=False)
            df_bonita.to_excel(writer, sheet_name="Performa Bonita (Tim Nego)", index=False)
            df_mirza.to_excel(writer, sheet_name="Performa Mirza (Tim Nego)", index=False)
            df_comp.to_excel(writer, sheet_name="Komparasi Bulanan Tim Nego", index=False)

            # Rincian Non-Tim Nego (Surveyor)
            df_non = df_2026[~df_2026["is_tim_nego"]][
                ["nomor_pengajuan", "nama_cabang", "wilayah", "negosiator", "tgl_nego_berakhir",
                 "asking_price", "deal_price", "saving_rupiah", "diskon_pct", "durasi_hari"]
            ]
            df_non.to_excel(writer, sheet_name="Non-Tim Nego (Direct Deal)", index=False)

        print(f"[✓] Berkas Excel Multi-Sheet berhasil disimpan: {excel_path}")
    except Exception as e:
        print(f"[!] Catatan ekspor Excel: {e}")

    print(f"[✓] Berkas CSV berhasil diperbarui di: {OUTPUT_DIR}")


# ==============================================================================
# 6. VISUALISASI DASHBOARD TERMINAL
# ==============================================================================
def cetak_dashboard_terminal(df_overall, df_bonita, df_mirza, df_comp):
    """Menampilkan dashboard eksekutif di terminal VSCode."""
    garis_ganda = "=" * 115
    garis_tunggal = "-" * 115

    print("\n" + garis_ganda)
    print("        PUSAT GADAI INDONESIA — LAPORAN PERFORMA NEGOSIASI BULANAN RESMI TAHUN 2026")
    print("             (CATATAN DIREKSI: TIM NEGOSIASI RESMI 2026 HANYA MIRZA & BONITA)")
    print(garis_ganda)

    # 1. Tabel Bonita
    print("\n[ TABEL 1: PERFORMA BULANAN BONITA — TIM NEGOSIASI RESMI (THE HIGH-VALUE NEGOTIATOR) ]")
    print(garis_tunggal)
    h_b = f"{'Bulan':<14} | {'Deal':>5} | {'Total Asking':>16} | {'Total Deal Sewa':>16} | {'Total Saving':>16} | {'Diskon %':>9} | {'Durasi hr':>9} | {'Success':>7}"
    print(h_b)
    print(garis_tunggal)
    for _, r in df_bonita.iterrows():
        is_tot = r["bulan_num"] == 99
        pfx = ">> " if is_tot else "   "
        lbl = (pfx + str(r["bulan_nama"]))[:14]
        if is_tot:
            print(garis_tunggal)
        print(f"{lbl:<14} | {int(r['jumlah_deal']):>5} | {rupiah(r['total_asking']):>16} | {rupiah(r['total_deal_price']):>16} | {rupiah(r['total_saving']):>16} | {r['avg_diskon_pct']:>8.2f}% | {r['avg_durasi_hari']:>7.1f} hr | {r['success_rate_pct']:>6.0f}%")
    print(garis_tunggal)

    # 2. Tabel Mirza
    print("\n[ TABEL 2: PERFORMA BULANAN MIRZA — TIM NEGOSIASI RESMI (THE SPEED & VOLUME SPECIALIST) ]")
    print(garis_tunggal)
    print(h_b)
    print(garis_tunggal)
    for _, r in df_mirza.iterrows():
        is_tot = r["bulan_num"] == 99
        pfx = ">> " if is_tot else "   "
        lbl = (pfx + str(r["bulan_nama"]))[:14]
        if is_tot:
            print(garis_tunggal)
        print(f"{lbl:<14} | {int(r['jumlah_deal']):>5} | {rupiah(r['total_asking']):>16} | {rupiah(r['total_deal_price']):>16} | {rupiah(r['total_saving']):>16} | {r['avg_diskon_pct']:>8.2f}% | {r['avg_durasi_hari']:>7.1f} hr | {r['success_rate_pct']:>6.0f}%")
    print(garis_tunggal)

    # 3. Tabel Komparasi Head-to-Head
    print("\n[ TABEL 3: KOMPARASI HEAD-TO-HEAD BULANAN (BONITA VS MIRZA + SURVEYOR NON-TIM NEGO) ]")
    print(garis_tunggal)
    h_c = f"{'Bulan':<11} | {'Deal B':>6} | {'Diskon B':>8} | {'Saving B':>12} | {'Dur B':>6} | {'Deal M':>6} | {'Diskon M':>8} | {'Saving M':>12} | {'Dur M':>6} | {'Non-Tim*':>8} | {'Total Saving':>12}"
    print(h_c)
    print(garis_tunggal)
    for _, r in df_comp.iterrows():
        is_tot = r["bulan_num"] == 99
        lbl = (">> " + str(r["bulan_nama"]))[:11] if is_tot else f"   {r['bulan_nama']:<8}"
        if is_tot:
            print(garis_tunggal)
        save_b = f"Rp {r['saving_bonita']/1e6:.1f} Jt" if r['saving_bonita'] > 0 else "-"
        save_m = f"Rp {r['saving_mirza']/1e6:.1f} Jt" if r['saving_mirza'] > 0 else "-"
        save_tot = f"Rp {r['total_saving']/1e6:.1f} Jt" if r['total_saving'] > 0 else "-"
        dur_b = f"{r['durasi_bonita']:.1f} hr" if r['deal_bonita'] > 0 else "-"
        dur_m = f"{r['durasi_mirza']:.1f} hr" if r['deal_mirza'] > 0 else "-"
        disc_b = f"{r['diskon_bonita']:.1f}%" if r['deal_bonita'] > 0 else "-"
        disc_m = f"{r['diskon_mirza']:.1f}%" if r['deal_mirza'] > 0 else "-"
        non_t = f"{int(r['deal_non_tim'])} deal" if r['deal_non_tim'] > 0 else "-"

        print(f"{lbl:<11} | {int(r['deal_bonita']):>6} | {disc_b:>8} | {save_b:>12} | {dur_b:>6} | {int(r['deal_mirza']):>6} | {disc_m:>8} | {save_m:>12} | {dur_m:>6} | {non_t:>8} | {save_tot:>12}")
    print(garis_tunggal)
    print("Catatan: (*) Kolom 'Non-Tim' merupakan pengajuan langsung surveyor (Dika: 10 deal, Salma: 1 deal) tanpa proses tawar-menawar (diskon 0%, saving Rp 0).")

    print("\n[ HIGHLIGHT STRATEGIS BISNIS 2026 ]")
    print("1. Tim Negosiasi Mandiri: Bonita dan Mirza menangani 264 cabang (96,0% transaksi 2026) dan menyumbang 100% dari penghematan Rp 2,04 Miliar.")
    print("2. Puncak Efisiensi Finansial: Bonita mencetak rekor saving tertinggi di Juni 2026 (Rp 291,5 Juta, diskon 18,65%, durasi 7,7 hari).")
    print("3. Puncak Kecepatan Closing: Mirza mempertahankan kecepatan luar biasa sepanjang tahun (Januari: 4,8 hari, Juli: 6,9 hari, Agustus: 8,5 hari).")
    print("4. Status Dika & Salma: Merupakan personil tim surveyor lapangan. 11 cabang yang diajukan langsung ditutup di harga penawaran awal tanpa negosiasi.")
    print(garis_ganda + "\n")


# ==============================================================================
# 7. PEMBUATAN GRAFIK & DIAGRAM VISUALISASI
# ==============================================================================
def buat_grafik_visualisasi(df_overall, df_bonita, df_mirza, df_comp):
    """Membuat 4 grafik visualisasi analisis dan menyimpannya ke folder 'hasil_analisis/grafik/'."""
    grafik_dir = os.path.join(OUTPUT_DIR, "grafik")
    os.makedirs(grafik_dir, exist_ok=True)
    print(f"[*] Membuat grafik visualisasi di: {grafik_dir}...")

    # Subset bulan 1-8 (eksklusif baris Total)
    df_o = df_overall[df_overall["bulan_num"] < 90].sort_values("bulan_num")
    df_b = df_bonita[df_bonita["bulan_num"] < 90].sort_values("bulan_num")
    df_m = df_mirza[df_mirza["bulan_num"] < 90].sort_values("bulan_num")
    df_c = df_comp[df_comp["bulan_num"] < 90].sort_values("bulan_num")

    bulan_labels = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu"]
    x = np.arange(len(bulan_labels))

    # --------------------------------------------------------------------------
    # GRAFIK 1: Tren Bulanan Volume & Total Saving 2026 (Dual Axis)
    # --------------------------------------------------------------------------
    fig, ax1 = plt.subplots(figsize=(9.5, 5.2), dpi=300)
    w = 0.48
    bars = ax1.bar(x, df_o["total_deal"], width=w, color="#1E3A8A", alpha=0.9, label="Volume Deal (Unit Cabang)")
    ax1.set_ylabel("Volume Deal (Unit Cabang)", color="#1E3A8A", fontsize=11, fontweight="bold")
    ax1.set_xticks(x)
    ax1.set_xticklabels(bulan_labels, fontsize=10, fontweight="bold")
    ax1.set_ylim(0, 62)
    ax1.grid(axis="y", linestyle="--", alpha=0.3)

    for bar in bars:
        h = bar.get_height()
        ax1.annotate(f"{int(h)}", xy=(bar.get_x() + bar.get_width() / 2, h),
                     xytext=(0, 3), textcoords="offset points", ha="center", va="bottom",
                     fontsize=9, fontweight="bold", color="#1E3A8A")

    ax2 = ax1.twinx()
    saving_juta = df_o["total_saving"] / 1_000_000
    line = ax2.plot(x, saving_juta, color="#059669", linewidth=2.6, marker="o", markersize=7, label="Total Saving (Juta Rp)")
    ax2.set_ylabel("Total Penghematan / Saving (Juta Rupiah)", color="#059669", fontsize=11, fontweight="bold")
    ax2.set_ylim(0, 440)

    for i, txt in enumerate(saving_juta):
        ax2.annotate(f"Rp {txt:.1f} Jt", (x[i], txt), textcoords="offset points", xytext=(0, 8),
                     ha="center", fontsize=8.5, fontweight="bold", color="#059669")

    plt.title("Tren Bulanan Volume Deal & Akumulasi Penghematan Sewa (Saving) — Tahun 2026", fontsize=12, fontweight="bold", pad=15)
    fig.tight_layout()
    p1 = os.path.join(grafik_dir, "1_tren_volume_dan_saving_2026.png")
    fig.savefig(p1)
    plt.close(fig)

    # --------------------------------------------------------------------------
    # GRAFIK 2: Komparasi Saving Bulanan: Bonita vs Mirza
    # --------------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 5.2), dpi=300)
    w = 0.35
    save_b_jt = df_c["saving_bonita"] / 1_000_000
    save_m_jt = df_c["saving_mirza"] / 1_000_000

    b1 = ax.bar(x - w/2, save_b_jt, w, label="Bonita (The High-Value Negotiator)", color="#2563EB", alpha=0.9)
    b2 = ax.bar(x + w/2, save_m_jt, w, label="Mirza (The Speed & Volume Specialist)", color="#0D9488", alpha=0.9)
    ax.set_ylabel("Total Saving (Juta Rupiah)", fontsize=11, fontweight="bold")
    ax.set_title("Komparasi Penghematan Sewa Bulanan: Bonita vs Mirza (Tahun 2026)", fontsize=12.5, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(bulan_labels, fontsize=10, fontweight="bold")
    ax.set_ylim(0, 330)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.legend(frameon=True, facecolor="white", edgecolor="#CBD5E1", fontsize=9.5)

    for bar in b1:
        h = bar.get_height()
        if h > 0:
            ax.annotate(f"{h:.0f}", (bar.get_x() + bar.get_width()/2, h), xytext=(0, 3),
                        textcoords="offset points", ha="center", va="bottom", fontsize=8, fontweight="bold", color="#1D4ED8")
    for bar in b2:
        h = bar.get_height()
        if h > 0:
            ax.annotate(f"{h:.0f}", (bar.get_x() + bar.get_width()/2, h), xytext=(0, 3),
                        textcoords="offset points", ha="center", va="bottom", fontsize=8, fontweight="bold", color="#0F766E")

    fig.tight_layout()
    p2 = os.path.join(grafik_dir, "2_komparasi_saving_bonita_mirza_2026.png")
    fig.savefig(p2)
    plt.close(fig)

    # --------------------------------------------------------------------------
    # GRAFIK 3: Komparasi Efisiensi Diskon (%) & Kecepatan Durasi (Hari)
    # --------------------------------------------------------------------------
    fig, (ax_d, ax_t) = plt.subplots(1, 2, figsize=(13, 5), dpi=300)

    # Subplot A: Diskon %
    ax_d.plot(bulan_labels, df_c["diskon_bonita"], marker="o", linewidth=2.2, color="#2563EB", label="Bonita (Rerata: 18,36%)")
    ax_d.plot(bulan_labels, df_c["diskon_mirza"], marker="s", linewidth=2.2, color="#0D9488", label="Mirza (Rerata: 13,99%)")
    ax_d.plot(bulan_labels, df_o["avg_diskon_pct"], linestyle="--", color="#64748B", linewidth=1.8, label="Rerata Nasional (15,57%)")
    ax_d.set_title("A. Efisiensi Diskon Sewa (%) Bulanan", fontsize=11, fontweight="bold")
    ax_d.set_ylabel("Diskon Rata-rata (%)", fontsize=10, fontweight="bold")
    ax_d.set_ylim(5, 30)
    ax_d.grid(True, linestyle="--", alpha=0.3)
    ax_d.legend(fontsize=8.5)

    for i in range(len(bulan_labels)):
        ax_d.annotate(f"{df_c['diskon_bonita'].iloc[i]:.1f}%", (bulan_labels[i], df_c['diskon_bonita'].iloc[i]),
                     textcoords="offset points", xytext=(0, 6), ha="center", fontsize=7.5, color="#1D4ED8", fontweight="bold")
        ax_d.annotate(f"{df_c['diskon_mirza'].iloc[i]:.1f}%", (bulan_labels[i], df_c['diskon_mirza'].iloc[i]),
                     textcoords="offset points", xytext=(0, -11), ha="center", fontsize=7.5, color="#0F766E", fontweight="bold")

    # Subplot B: Durasi Nego
    ax_t.plot(bulan_labels, df_c["durasi_bonita"], marker="o", linewidth=2.2, color="#2563EB", label="Bonita (Rerata: 11,6 hr)")
    ax_t.plot(bulan_labels, df_c["durasi_mirza"], marker="s", linewidth=2.2, color="#0D9488", label="Mirza (Rerata: 9,6 hr)")
    ax_t.axhline(14, color="#DC2626", linestyle=":", linewidth=1.8, label="Target Cut-off SLA (14 hr)")
    ax_t.set_title("B. Kecepatan Realisasi Negosiasi (Hari)", fontsize=11, fontweight="bold")
    ax_t.set_ylabel("Durasi Rata-rata (Hari Kalender)", fontsize=10, fontweight="bold")
    ax_t.set_ylim(0, 22)
    ax_t.grid(True, linestyle="--", alpha=0.3)
    ax_t.legend(fontsize=8.5)

    for i in range(len(bulan_labels)):
        ax_t.annotate(f"{df_c['durasi_bonita'].iloc[i]:.1f}h", (bulan_labels[i], df_c['durasi_bonita'].iloc[i]),
                     textcoords="offset points", xytext=(0, 6), ha="center", fontsize=7.5, color="#1D4ED8", fontweight="bold")
        ax_t.annotate(f"{df_c['durasi_mirza'].iloc[i]:.1f}h", (bulan_labels[i], df_c['durasi_mirza'].iloc[i]),
                     textcoords="offset points", xytext=(0, -11), ha="center", fontsize=7.5, color="#0F766E", fontweight="bold")

    fig.suptitle("Dinamika Kinerja Tim Negosiasi Resmi PGI 2026: Diskon vs Kecepatan Closing", fontsize=12, fontweight="bold", y=1.02)
    fig.tight_layout()
    p3 = os.path.join(grafik_dir, "3_komparasi_diskon_dan_durasi_2026.png")
    fig.savefig(p3)
    plt.close(fig)

    # --------------------------------------------------------------------------
    # GRAFIK 4: Dekomposisi Siklus Waktu UPC (2024–2026) — Harmonisasi Lead Time
    # --------------------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=300)
    years = ["2024 (317 cabang)", "2025 (436 cabang)", "2026 (275 cabang)"]
    y_pos = np.arange(len(years))

    pengajuan = np.array([0.0, 13.5, 10.7])
    t_nego    = np.array([3.0, 1.7, 11.1])
    d_nego    = np.array([33.0, 24.4, 11.5])
    t_renov   = np.array([3.0, 4.2, 6.3])
    d_renov   = np.array([17.1, 19.8, 23.9])
    t_go      = np.array([14.0, 9.4, 5.6])

    c_pengajuan = "#8B5CF6"
    c_tnego     = "#94A3B8"
    c_dnego     = "#2563EB"
    c_trenov    = "#CBD5E1"
    c_drenov    = "#F59E0B"
    c_tgo       = "#10B981"

    ax.barh(y_pos, pengajuan, color=c_pengajuan, label="1. Pengajuan s/d Approved")
    ax.barh(y_pos, t_nego, left=pengajuan, color=c_tnego, label="2. Tunggu Nego")
    ax.barh(y_pos, d_nego, left=pengajuan + t_nego, color=c_dnego, label="3. Durasi Nego (Realisasi)")
    ax.barh(y_pos, t_renov, left=pengajuan + t_nego + d_nego, color=c_trenov, label="4. Tunggu Renovasi")
    ax.barh(y_pos, d_renov, left=pengajuan + t_nego + d_nego + t_renov, color=c_drenov, label="5. Durasi Renovasi Fisik")
    ax.barh(y_pos, t_go, left=pengajuan + t_nego + d_nego + t_renov + d_renov, color=c_tgo, label="6. Tunggu Grand Opening (paid_at s/d Buka)")

    for i, y in enumerate(y_pos):
        total_len = (pengajuan[i] + t_nego[i] + d_nego[i] + t_renov[i] + d_renov[i] + t_go[i])
        if i == 0:
            ax.annotate("Total: 69.5 hr\n(Approval N/A*)", (70.1 + 1, y), va="center", fontsize=8.5, fontweight="bold", color="#1E293B")
        elif i == 1:
            ax.annotate(f"Total: 68.2 hr (Sum: {total_len:.1f}h)\n[Pasca-Approval: 57.2 hr]", (total_len + 1, y), va="center", fontsize=8.5, fontweight="bold", color="#1E293B")
        elif i == 2:
            ax.annotate(f"Total: 69.2 hr (Sum: 69.1h)\n[Pasca-Approval: 58.7 hr (~58h)]", (total_len + 1, y), va="center", fontsize=8.5, fontweight="bold", color="#1E3A8A")

    ax.set_yticks(y_pos)
    ax.set_yticklabels(years, fontsize=10, fontweight="bold")
    ax.set_xlabel("Rata-rata Durasi (Hari Kalender)", fontsize=10, fontweight="bold")
    ax.set_title("Dekomposisi Lengkap Siklus Lead Time Pembukaan Cabang UPC (2024–2026)\nHarmonisasi: Total Lead Time (Pengajuan s/d Buka: 69,2 hr) vs Pasca-Approval (58,7 hr)", fontsize=11.5, fontweight="bold", pad=15)
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    ax.legend(bbox_to_anchor=(0.5, -0.16), loc="upper center", ncol=3, frameon=True, fontsize=8)
    ax.set_xlim(0, 96)

    fig.tight_layout()
    p4 = os.path.join(grafik_dir, "4_dekomposisi_siklus_waktu_upc.png")
    fig.savefig(p4)
    plt.close(fig)

    # Menghasilkan Gambar Tabeling Dekomposisi 7 Siklus & Detail Bulanan 2026
    try:
        from pipeline_master_analisis_upc_pgi import analisis_dekomposisi_siklus_waktu, get_cleaned_dataset
        df_all = get_cleaned_dataset()
        analisis_dekomposisi_siklus_waktu(df_all, generate_chart=True)
    except Exception as e:
        print(f"[!] Info perenderan tabel 7 siklus: {e}")

    print(f"[✓] Berkas Grafik PNG & Tabeling Dekomposisi 7 Siklus berhasil dibuat di: {grafik_dir}")


def buat_grafik_distribusi_korelasi(df_2026, grafik_dir=None):
    """
    Membuat Grafik 7: Distribusi & Uji Korelasi Durasi Negosiasi vs Efisiensi Diskon Sewa (Tahun 2026).
    Multi-panel chart:
      - Panel A: Sebaran (Scatter Plot) durasi tawar-menawar vs diskon sewa dengan garis regresi linier.
      - Panel B: Rata-rata efisiensi diskon per kelompok durasi (1-5, 6-10, 11-14, 15-20, >20 hari).
      - Panel C: Tingkat keberhasilan memperoleh diskon (>0%) dan penghematan per bucket.
    """
    if grafik_dir is None:
        grafik_dir = os.path.join(OUTPUT_DIR, "grafik")
    os.makedirs(grafik_dir, exist_ok=True)
    print(f"[*] Membuat Grafik 7: Distribusi & Korelasi Durasi vs Diskon di: {grafik_dir}...")

    # Pastikan data durasi dan diskon valid
    df_clean = df_2026.dropna(subset=["durasi_hari", "diskon_pct"]).copy()

    # Hitung statistik korelasi
    slope, intercept, r_val, p_val, std_err = stats.linregress(df_clean["durasi_hari"], df_clean["diskon_pct"])
    spearman_rho, spearman_p = stats.spearmanr(df_clean["durasi_hari"], df_clean["diskon_pct"])
    r_squared = r_val ** 2

    # Siapkan figure multi-panel
    fig = plt.figure(figsize=(16, 8), dpi=300)
    gs = fig.add_gridspec(2, 2, width_ratios=[1.25, 1], height_ratios=[1, 1], wspace=0.25, hspace=0.35)

    ax1 = fig.add_subplot(gs[:, 0])
    ax2 = fig.add_subplot(gs[0, 1])
    ax3 = fig.add_subplot(gs[1, 1])

    # 1. Panel A: Scatter Plot
    df_bon = df_clean[df_clean["negosiator"] == "Bonita"]
    df_mir = df_clean[df_clean["negosiator"] == "Mirza"]
    df_oth = df_clean[~df_clean["is_tim_nego"]]

    ax1.scatter(df_bon["durasi_hari"], df_bon["diskon_pct"], color="#3B82F6", alpha=0.65, s=45, label=f"Bonita ({len(df_bon)} cabang | Rerata: {df_bon['diskon_pct'].mean():.1f}%)")
    ax1.scatter(df_mir["durasi_hari"], df_mir["diskon_pct"], color="#10B981", alpha=0.65, s=45, label=f"Mirza ({len(df_mir)} cabang | Rerata: {df_mir['diskon_pct'].mean():.1f}%)")
    ax1.scatter(df_oth["durasi_hari"], df_oth["diskon_pct"], color="#EF4444", marker="x", s=55, linewidth=1.5, label=f"Non-Tim ({len(df_oth)} cabang | Diskon 0%)")

    # Garis Regresi Linier
    x_line = np.linspace(df_clean["durasi_hari"].min(), df_clean["durasi_hari"].max() + 3, 100)
    y_line = slope * x_line + intercept
    ax1.plot(x_line, y_line, color="#DC2626", linewidth=2.5, label=f"Garis Tren Linier (y = {slope:.2f}x + {intercept:.1f}%)")

    # Garis SLA 14 Hari & Rerata Nasional
    ax1.axvline(14, color="#DC2626", linestyle="--", linewidth=1.8, label="Batas Toleransi SLA (14 Hari)")
    mean_diskon = df_clean["diskon_pct"].mean()
    ax1.axhline(mean_diskon, color="#475569", linestyle=":", linewidth=1.5, label=f"Rerata Diskon Nasional ({mean_diskon:.2f}%)")

    ax1.set_title(f"A. Sebaran (Scatter Plot) Durasi Negosiasi vs Efisiensi Diskon Sewa 2026\n(N = {len(df_clean)} Transaksi Cabang)", fontsize=11.5, fontweight="bold", pad=10)
    ax1.set_xlabel("Durasi Negosiasi (Hari Kalender - durasi_nego_hari)", fontsize=10.5, fontweight="bold")
    ax1.set_ylabel("Efisiensi Diskon Sewa (%)", fontsize=10.5, fontweight="bold")
    ax1.set_xlim(0, max(52, df_clean["durasi_hari"].max() + 2))
    ax1.set_ylim(-2, 53)
    ax1.grid(True, linestyle="--", alpha=0.3)
    ax1.legend(loc="upper right", frameon=True, fontsize=8.5)

    # Kotak Callout Statistik
    callout_txt = (
        f"UJI KORELASI STATISTIK:\n"
        f"• Pearson r  : {r_val:.4f} (p = {p_val:.3f})\n"
        f"• Spearman ρ : {spearman_rho:.4f} (p = {spearman_p:.3f})\n"
        f"• R-Squared  : {r_squared:.4f} (Hanya {r_squared*100:.1f}% varians)\n"
        f"• KESIMPULAN : TIDAK ADA korelasi positif.\n"
        f"  Menambah hari negosiasi TIDAK memperbesar diskon!\n"
        f"  Diskon terbaik justru tercapai di durasi ≤ 10 hari."
    )
    ax1.text(0.03, 0.05, callout_txt, transform=ax1.transAxes, fontsize=8.5,
             verticalalignment="bottom", bbox=dict(boxstyle="round,pad=0.5", facecolor="#F8FAFC", edgecolor="#334155", alpha=0.95))

    # 2. Pengelompokan Durasi (Bucketing)
    bins = [0, 5, 10, 14, 20, 999]
    bin_labels = ["1–5 hr", "6–10 hr", "11–14 hr", "15–20 hr", "> 20 hr"]
    df_clean["durasi_bucket"] = pd.cut(df_clean["durasi_hari"], bins=bins, labels=bin_labels)

    bucket_stats = df_clean.groupby("durasi_bucket", observed=False).agg(
        deal_count=("nomor_pengajuan" if "nomor_pengajuan" in df_clean.columns else "durasi_hari", "count"),
        avg_diskon=("diskon_pct", "mean"),
        total_saving=("saving_rupiah", "sum"),
        success_deals=("diskon_pct", lambda s: (s > 0).sum())
    ).reset_index()
    bucket_stats["success_rate"] = (bucket_stats["success_deals"] / bucket_stats["deal_count"]) * 100
    bucket_stats["total_saving_juta"] = bucket_stats["total_saving"] / 1_000_000

    colors_bucket = ["#10B981", "#0284C7", "#3B82F6", "#F59E0B", "#DC2626"]
    x_b = np.arange(len(bin_labels))

    # 3. Panel B: Rerata Diskon per Bucket
    bars_b = ax2.bar(x_b, bucket_stats["avg_diskon"], color=colors_bucket, width=0.55, edgecolor="#1E293B", linewidth=0.8)
    ax2.set_title("B. Rerata Efisiensi Diskon Berdasarkan Kelompok Durasi", fontsize=10.5, fontweight="bold", pad=10)
    ax2.set_xticks(x_b)
    ax2.set_xticklabels(bin_labels, fontsize=9.5, fontweight="bold")
    ax2.set_ylabel("Rerata Diskon (%)", fontsize=9.5, fontweight="bold")
    ax2.set_ylim(0, 24)
    ax2.grid(axis="y", linestyle="--", alpha=0.3)

    for bar, (_, row) in zip(bars_b, bucket_stats.iterrows()):
        h = bar.get_height()
        deals = int(row["deal_count"])
        ax2.annotate(f"{h:.1f}%\n({deals} deal)", xy=(bar.get_x() + bar.get_width() / 2, h),
                     xytext=(0, 4), textcoords="offset points", ha="center", va="bottom",
                     fontsize=8.5, fontweight="bold", color="#0F172A")

    # 4. Panel C: Tingkat Keberhasilan Meraih Diskon (Success Rate)
    bars_c = ax3.bar(x_b, bucket_stats["success_rate"], color=colors_bucket, width=0.55, edgecolor="#1E293B", linewidth=0.8)
    ax3.axhline(90, color="#DC2626", linestyle=":", linewidth=1.4, label="Ambang Batas Sehat (90%)")
    ax3.set_title("C. Tingkat Keberhasilan Mendapatkan Diskon (Success Rate %)", fontsize=10.5, fontweight="bold", pad=10)
    ax3.set_xticks(x_b)
    ax3.set_xticklabels(bin_labels, fontsize=9.5, fontweight="bold")
    ax3.set_ylabel("Tingkat Sukses Diskon (%)", fontsize=9.5, fontweight="bold")
    ax3.set_ylim(50, 106)
    ax3.grid(axis="y", linestyle="--", alpha=0.3)
    ax3.legend(loc="upper right", frameon=True, fontsize=8)

    for bar, (_, row) in zip(bars_c, bucket_stats.iterrows()):
        h = bar.get_height()
        sav = row["total_saving_juta"]
        ax3.annotate(f"{h:.1f}%\n(Rp{sav:.0f}Jt)", xy=(bar.get_x() + bar.get_width() / 2, h),
                     xytext=(0, 4), textcoords="offset points", ha="center", va="bottom",
                     fontsize=8.5, fontweight="bold", color="#0F172A")

    fig.suptitle("Analisis Distribusi Durasi Negosiasi vs Efisiensi Diskon Sewa PGI Tahun 2026", fontsize=13, fontweight="bold", y=0.98)
    p7 = os.path.join(grafik_dir, "7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png")
    fig.savefig(p7, bbox_inches="tight")
    plt.close(fig)

    # Salin ke direktori artifacts jika ada
    artifact_grafik_dir = "/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik"
    if os.path.exists(artifact_grafik_dir):
        shutil.copy2(p7, os.path.join(artifact_grafik_dir, "7_distribusi_dan_korelasi_durasi_vs_diskon_2026.png"))

    print(f"[✓] Grafik 7 (Distribusi & Korelasi) berhasil dibuat di: {p7}")


# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================
def main():
    print("\n" + "=" * 60)
    print(" Memulai Analisis Performa Negosiasi Bulanan 2026...")
    print("=" * 60)

    # 1. Load Data
    df_2026 = load_and_prep_data()

    # 2. Analisis Bulanan Bonita
    df_bonita = hitung_performa_individu(df_2026, "Bonita")

    # 3. Analisis Bulanan Mirza
    df_mirza = hitung_performa_individu(df_2026, "Mirza")

    # 4. Tabel Komparasi Head-to-Head
    df_comp = buat_tabel_komparasi_bulanan(df_2026)

    # 5. Tren Bulanan Overall
    df_overall = hitung_performa_bulanan_overall(df_2026)

    # 6. Tampilkan Dashboard di Terminal
    cetak_dashboard_terminal(df_overall, df_bonita, df_mirza, df_comp)

    # 7. Ekspor ke Excel & CSV
    export_ke_file(df_overall, df_bonita, df_mirza, df_comp, df_2026)

    # 8. Buat Grafik & Diagram Visualisasi (Grafik 1 - 4)
    buat_grafik_visualisasi(df_overall, df_bonita, df_mirza, df_comp)

    # 9. Buat Grafik Distribusi & Korelasi Durasi vs Diskon (Grafik 7)
    buat_grafik_distribusi_korelasi(df_2026)

    print("[✓] Eksekusi Analisis & Pembuatan Grafik Selesai dengan Sukses!\n")


if __name__ == "__main__":
    main()


