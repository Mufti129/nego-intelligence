#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================================================
PIPELINE PEMBERSIHAN DATA (DATA CLEANSING & HARMONISASI MULTI-SUMBER) UPC PGI
====================================================================================================
Penulis      : Data Analyst — Divisi Bisnis (Pusat Gadai Indonesia)
Sumber Data  : Folder data_nego_baru/
               1. data_nego_baru_fix_open(jan-aug26).xlsx (Sheet 'data')
               2. Laporan Negosiator.xlsx (Sheet 'Bonita Baru' & 'Mirza Baru')
               3. tgl_open_cabang.xlsx (Sheet 'tgl_op')
Output File  : properties_cleaned_2024-2026.csv (Root directory)
               data_nego_baru/data_nego_baru_cleaned_2024-2026.xlsx (Excel siap audit)

Deskripsi Logika Bisnis & Perhitungan Durasi (7 Siklus Bersih & Ramping):
  1. lama_waktu_realisasi_pengajuan_ke_aprooved :
     Durasi sejak dokumen pengajuan dibuat hingga disetujui (approved_at - application_date).
  2. waktu_tunggu_aproved_ke_tgl_awal_nego :
     Waktu jeda sejak disetujui hingga mulai negosiasi (tanggal_awal_nego - approved_at).
  3. durasi_nego_hari :
     Durasi negosiasi aktif dari awal hingga deal (tgl_nego_berakhir - tanggal_awal_nego).
  4. lama_waktu_pengumpulan_berkas :
     Durasi pengumpulan berkas dari deal hingga ttd sewa (tgl_ttd_sewa - tgl_nego_berakhir).
  5. waktu_tunggu_sewa_ke_renovasi_awal :
     Waktu jeda sejak ttd sewa hingga renovasi dimulai (tgl_awal_renovasi - tgl_ttd_sewa).
  6. lama_waktu_realisasi_renovasi_Selesai :
     Durasi aktual pekerjaan fisik renovasi (tgl_realisasi_akhir_renovasi - tgl_awal_renovasi).
  7. waktu_tunggu_selesai_renove_open_cabang :
     Waktu jeda persiapan akhir renovasi selesai s/d grand opening (Open_cabang - tgl_realisasi_akhir_renovasi).
  8. lama_proses_pembukaan_cabang :
     Total lead time pembukaan cabang end-to-end (Open_cabang - application_date).
  9. sla_renov :
     Target SLA waktu renovasi dari manajemen (target_tgl_akhir_renovasi - tgl_awal_renovasi).
 10. sla_complimence_persen :
     Status "tercapai" jika durasi aktual <= target SLA, sebaliknya "tidak tercapai".
====================================================================================================
"""

import os
import sys
import numpy as np
import pandas as pd


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_NEGO_DIR = os.path.join(BASE_DIR, "data_nego_baru")

FILE_MAIN_EXCEL = os.path.join(DATA_NEGO_DIR, "data_nego_baru_fix_open(jan-aug26).xlsx")
FILE_LAPORAN_NEGO = os.path.join(DATA_NEGO_DIR, "Laporan Negosiator.xlsx")
FILE_TGL_OPEN = os.path.join(DATA_NEGO_DIR, "tgl_open_cabang.xlsx")

OUTPUT_CSV = os.path.join(BASE_DIR, "properties_cleaned_2024-2026.csv")
OUTPUT_EXCEL = os.path.join(DATA_NEGO_DIR, "data_nego_baru_cleaned_2024-2026.xlsx")


def rupiah(nilai):
    """Format angka ke rupiah untuk tampilan audit."""
    if pd.isna(nilai) or nilai == 0:
        return "Rp 0"
    abs_val = abs(nilai)
    if abs_val >= 1_000_000_000:
        return f"Rp {nilai / 1_000_000_000:.2f} Miliar"
    elif abs_val >= 1_000_000:
        return f"Rp {nilai / 1_000_000:.1f} Juta"
    else:
        return f"Rp {nilai:,.0f}".replace(",", ".")


def run_cleansing(silent=False):
    if not silent:
        print("=" * 95)
        print("      MEMULAI PIPELINE PEMBERSIHAN DATA (DATA CLEANSING & HARMONISASI BARU)")
        print("=" * 95)

    # 1. Validasi Berkas Masukan
    for fpath in [FILE_MAIN_EXCEL, FILE_LAPORAN_NEGO, FILE_TGL_OPEN]:
        if not os.path.exists(fpath):
            print(f"[!] ERROR: Berkas tidak ditemukan: {fpath}")
            sys.exit(1)

    if not silent:
        print(f"[*] 1. Membaca dataset utama: {os.path.basename(FILE_MAIN_EXCEL)} (sheet: 'data')")
    df_raw = pd.read_excel(FILE_MAIN_EXCEL, sheet_name="data")
    total_raw = len(df_raw)
    if not silent:
        print(f"       -> Total baris mentah : {total_raw} baris")
        print(f"       -> Cabang unik awal   : {df_raw['nomor_pengajuan'].nunique()} cabang (kolom E)")

    # Filter out cabang Pelebaran (karena bukan unit pembukaan cabang baru / UPC murni)
    mask_pelebaran = df_raw["nama_cabang"].astype(str).str.contains(r"pelebaran", case=False, na=False)
    total_pelebaran_raw = mask_pelebaran.sum()
    df_raw = df_raw[~mask_pelebaran].copy()
    if not silent:
        print(f"       -> Mengeluarkan {total_pelebaran_raw} baris bertanda '(Pelebaran)' (bukan unit baru UPC).")
        print(f"       -> Baris mentah setelah filter pelebaran : {len(df_raw)} baris")

    # 2. Baca Sheet Laporan Negosiator (Bonita Baru & Mirza Baru)
    if not silent:
        print(f"[*] 2. Membaca lembar kerja negosiator: {os.path.basename(FILE_LAPORAN_NEGO)}")
    df_bonita = pd.read_excel(FILE_LAPORAN_NEGO, sheet_name="Bonita Baru").dropna(subset=["ID Ruko"])
    df_mirza = pd.read_excel(FILE_LAPORAN_NEGO, sheet_name="Mirza Baru").dropna(subset=["Column 1"])

    df_bonita["key"] = df_bonita["ID Ruko"].astype(str).str.strip()
    df_mirza["key"] = df_mirza["Column 1"].astype(str).str.strip()
    df_raw["key"] = df_raw["nomor_pengajuan"].astype(str).str.strip()

    b_lookup = df_bonita.set_index("key")[["Harga Awal", "Harga Deal", "Efisiensi", "Persentase Efisiensi", "Negosiator"]].to_dict("index")
    m_lookup = df_mirza.set_index("key")[["Harga Awal", "Harga Deal", "Efisiensi", "Persentase Efisiensi", "Negosiator"]].to_dict("index")
    if not silent:
        print(f"       -> Data valid Bonita Baru : {len(b_lookup)} entri ruko")
        print(f"       -> Data valid Mirza Baru  : {len(m_lookup)} entri ruko")

    # 3. Baca Pemetaan Tanggal Buka (tgl_open_cabang.xlsx -> sheet 'tgl_op')
    if not silent:
        print(f"[*] 3. Membaca tanggal buka gerai: {os.path.basename(FILE_TGL_OPEN)} (sheet: 'tgl_op')")
    df_op = pd.read_excel(FILE_TGL_OPEN, sheet_name="tgl_op")
    df_op["name_clean"] = df_op["name"].astype(str).str.strip()
    op_map = dict(zip(df_op["name_clean"], pd.to_datetime(df_op["operational_date"], errors="coerce")))

    # Standardisasi nama cabang untuk pencocokan ke tgl_op
    clean_cabang = df_raw["nama_cabang"].astype(str).str.replace(r"\s*\(.*?\)", "", regex=True).str.strip().replace({"PDR0006": "PDR006"})
    df_raw["Open_cabang"] = clean_cabang.map(op_map)
    matched_open = df_raw["Open_cabang"].notna().sum()
    if not silent:
        print(f"       -> Pemetaan Open_cabang sukses: {matched_open} / {total_raw} baris")

    # 4. Deduplikasi Multi-Termin (keep last completion date)
    if not silent:
        print("[*] 4. Rekonsiliasi & deduplikasi pembayaran multi-termin...")
    df_sorted = df_raw.sort_values(by=["nomor_pengajuan", "tgl_realisasi_akhir_renovasi"])
    dup_count = df_sorted.duplicated(subset=["nomor_pengajuan"]).sum()
    df = df_sorted.drop_duplicates(subset=["nomor_pengajuan"], keep="last").copy()
    if not silent:
        print(f"       -> Baris multi-termin direkonsiliasi: {dup_count} baris")
        print(f"       -> Total cabang unik valid          : {len(df)} unit")

    # 5. Parsing & Validasi Tanggal
    # 5. Parsing & Validasi Tanggal (Dinormalisasi ke Hari Kalender Murni 00:00:00)
    t_app = pd.to_datetime(df["application_date"], errors="coerce").dt.normalize()
    t_appr = pd.to_datetime(df["approved_at"], errors="coerce").dt.normalize()
    t_nego_start = pd.to_datetime(df["tanggal_awal_nego"], errors="coerce").dt.normalize()
    t_nego_end = pd.to_datetime(df["tgl_nego_berakhir"], errors="coerce").dt.normalize()
    t_ttd = pd.to_datetime(df["tgl_ttd_sewa"], errors="coerce").dt.normalize()
    t_renov_start = pd.to_datetime(df["tgl_awal_renovasi"], errors="coerce").dt.normalize()
    t_renov_target = pd.to_datetime(df["target_tgl_akhir_renovasi"], errors="coerce").dt.normalize()
    t_renov_end = pd.to_datetime(df["tgl_realisasi_akhir_renovasi"], errors="coerce").dt.normalize()
    t_open = pd.to_datetime(df["Open_cabang"], errors="coerce").dt.normalize()

    # Koreksi anomali typo tahun pada tgl_nego_berakhir untuk data 2024
    typo_mask = (t_nego_end > t_ttd) & (t_app.dt.year < 2026)
    if typo_mask.sum() > 0:
        if not silent:
            print(f"       -> Menyelaraskan {typo_mask.sum()} baris anomali typo tahun tgl_nego_berakhir.")
        t_nego_end[typo_mask] = t_ttd[typo_mask]

    # Simpan kembali tanggal-tanggal yang telah diselaraskan ke dataframe
    df["application_date"] = t_app
    df["approved_at"] = t_appr
    df["tanggal_awal_nego"] = t_nego_start
    df["tgl_nego_berakhir"] = t_nego_end
    df["tgl_ttd_sewa"] = t_ttd
    df["tgl_awal_renovasi"] = t_renov_start
    df["target_tgl_akhir_renovasi"] = t_renov_target
    df["tgl_realisasi_akhir_renovasi"] = t_renov_end

    # 6. Perhitungan Ulang 10 Metrik Durasi & SLA Sesuai Logika Bisnis (7 Siklus Murni Kalender)
    if not silent:
        print("[*] 5. Menghitung ulang seluruh metrik durasi, lead time, dan kepatuhan SLA...")

    # 1. lama_waktu_realisasi_pengajuan_ke_aprooved (Siklus 1: Pengajuan ke Approval)
    d1 = (t_appr - t_app).dt.days
    df["lama_waktu_realisasi_pengajuan_ke_aprooved"] = d1.mask(d1 < 0)

    # 2. waktu_tunggu_aproved_ke_tgl_awal_nego (Siklus 2: Approval ke Awal Nego - Opsi A: Floor at 0)
    d2 = (t_nego_start - t_appr).dt.days
    df["waktu_tunggu_aproved_ke_tgl_awal_nego"] = np.maximum(0, d2)

    # 3. durasi_nego_hari (Siklus 3: Durasi Negosiasi Riil)
    d3 = (t_nego_end - t_nego_start).dt.days
    df["durasi_nego_hari"] = np.maximum(1, d3)

    # 4. lama_waktu_pengumpulan_berkas (Siklus 4: Pengumpulan Berkas / Deal ke TTD Sewa)
    d4 = (t_ttd - t_nego_end).dt.days
    df["lama_waktu_pengumpulan_berkas"] = np.maximum(0, d4)

    # 5. waktu_tunggu_sewa_ke_renovasi_awal (Siklus 5: TTD Sewa ke Awal Renovasi)
    d5 = (t_renov_start - t_ttd).dt.days
    df["waktu_tunggu_sewa_ke_renovasi_awal"] = np.maximum(0, d5)

    # 6. lama_waktu_realisasi_renovasi_Selesai (Siklus 6: Durasi Fisik Renovasi)
    d6 = (t_renov_end - t_renov_start).dt.days
    df["lama_waktu_realisasi_renovasi_Selesai"] = np.maximum(1, d6)

    # 7. waktu_tunggu_selesai_renove_open_cabang
    d7 = (t_open - t_renov_end).dt.days
    df["waktu_tunggu_selesai_renove_open_cabang"] = d7.mask(d7 < 0)

    # 8. lama_proses_pembukaan_cabang (Total End-to-End)
    d_total = (t_open - t_app).dt.days
    df["lama_proses_pembukaan_cabang"] = d_total.mask(d_total < 0)

    # 9. sla_renov
    sla_r = (t_renov_target - t_renov_start).dt.days
    df["sla_renov"] = sla_r

    # 10. sla_complimence_persen
    df["sla_complimence_persen"] = np.where(
        df["lama_waktu_realisasi_renovasi_Selesai"] <= df["sla_renov"],
        "tercapai",
        "tidak tercapai"
    )

    # 7. Integrasi Finansial & Personil dari Sheet Laporan Negosiator
    if not silent:
        print("[*] 6. Mengintegrasikan Harga Awal, Deal, Diskon, dan Negosiator dari Laporan Negosiator...")
    h_awal, h_deal, eff_nom, eff_pct, nego_2 = [], [], [], [], []

    for _, r in df.iterrows():
        k = r["key"]
        matched = False

        if k in b_lookup and k in m_lookup:
            # Jika ruko ada di kedua sheet, cek nama negosiator di df
            if str(r.get("nama_negosiator", "")).strip().lower() == "mirza":
                d = m_lookup[k]
            else:
                d = b_lookup[k]
            matched = True
        elif k in b_lookup:
            d = b_lookup[k]
            matched = True
        elif k in m_lookup:
            d = m_lookup[k]
            matched = True

        if matched:
            val_awal = float(d["Harga Awal"]) if pd.notnull(d["Harga Awal"]) else np.nan
            val_deal = float(d["Harga Deal"]) if pd.notnull(d["Harga Deal"]) else np.nan
            val_saving = float(d["Efisiensi"]) if pd.notnull(d["Efisiensi"]) else np.nan
            p_val = d["Persentase Efisiensi"]

            # Standardisasi Formula Diskon Kanonikal: Efisiensi / Harga Awal
            if pd.notnull(val_awal) and val_awal > 0 and pd.notnull(val_saving):
                val_pct = (val_saving / val_awal) * 100.0
            elif pd.notnull(p_val):
                val_pct = float(p_val) * 100.0 if float(p_val) <= 1.5 else float(p_val)
            else:
                val_pct = 0.0

            h_awal.append(val_awal)
            h_deal.append(val_deal)
            eff_nom.append(val_saving)
            eff_pct.append(val_pct)
            nego_2.append(str(d["Negosiator"]).strip().title())
        else:
            # Fallback untuk tahun 2024-2025 atau non-tim
            ask = float(r["asking_price"]) if pd.notnull(r.get("asking_price")) else np.nan
            rent = float(r["rental_price"]) if pd.notnull(r.get("rental_price")) else np.nan
            saving = max(0.0, ask - rent) if pd.notnull(ask) and pd.notnull(rent) else 0.0
            pct = (saving / ask * 100.0) if pd.notnull(ask) and ask > 0 else 0.0

            h_awal.append(ask)
            h_deal.append(rent)
            eff_nom.append(saving)
            eff_pct.append(pct)
            nego_2.append(str(r.get("nama_negosiator", "")).strip().title())

    df["harga_awal_penawaran"] = h_awal
    df["harga_rental_final"] = h_deal
    df["diskon_rupiah"] = eff_nom
    df["efisiensi_diskon_pct"] = eff_pct
    df["nama_negosiator_2"] = nego_2

    # 8. Periode Tahun, Bulan, dan negosiator_analisis
    df["Tahun"] = t_app.dt.year.fillna(t_nego_start.dt.year).astype(int)
    month_names = {
        1: "January", 2: "February", 3: "March", 4: "April",
        5: "May", 6: "June", 7: "July", 8: "August",
        9: "September", 10: "October", 11: "November", 12: "December"
    }
    df["Bulan"] = t_app.dt.month.map(month_names).fillna("Unknown") + " " + df["Tahun"].astype(str)

    df["negosiator_analisis"] = np.where(
        df["Tahun"] == 2026,
        df["nama_negosiator_2"].replace(["", "Nan", "None"], np.nan).fillna(df["nama_negosiator"]),
        df["nama_negosiator"]
    )

    # 9. Penyimpanan Hasil Cleansing ke CSV dan Excel (Ramping & Tanpa Duplikasi)
    if not silent:
        print("[*] 7. Menyimpan berkas hasil pembersihan yang telah dirampingkan...")
    COLS_TO_DROP = [
        "key",
        "sla_pure_nego_hari",
        "lama_waktu_realisasi_nego",
        "lama_waktu_pengumpulan berkas",
        "waktu_tunggu_nego_ke_renovasi_awal",
        "paid_at/tgl_realisasi_selesai_renovasi",
        "Hargaawal_nego_sheet",
        "hargarental_sheet",
        "efisiensi_diskon_nego",
        "efisiensi_diskon_nego2",
    ]
    df.drop(columns=COLS_TO_DROP, inplace=True, errors="ignore")
    df.to_csv(OUTPUT_CSV, index=False)
    if not silent:
        print(f"       [✓] Berkas CSV disimpan ke   : {OUTPUT_CSV}")

    try:
        df.to_excel(OUTPUT_EXCEL, index=False)
        if not silent:
            print(f"       [✓] Berkas Excel disimpan ke : {OUTPUT_EXCEL}")
    except Exception as e:
        if not silent:
            print(f"       [!] Gagal menyimpan excel: {e}")

    # 10. Laporan Ringkasan Hasil Cleansing
    if not silent:
        print("\n" + "=" * 95)
        print("                     RINGKASAN EKSEKUTIF DATA CLEANSING SELESAI")
        print("=" * 95)
        print(f"• Total Baris Mentah        : {total_raw} baris")
        print(f"• Total Cabang Unik Valid   : {len(df)} cabang (setelah deduplikasi multi-termin)")
        print(f"• Total Kolom Tersedia      : {len(df.columns)} kolom lengkap")
        print(f"• Cakupan Wilayah Indonesia : {df['wilayah'].nunique()} Kabupaten/Kota")

        print("\nDISTRIBUSI KOHORT TAHUNAN:")
        t_counts = df["Tahun"].value_counts().sort_index()
        for yr, cnt in t_counts.items():
            print(f"   - Kohort {yr} : {cnt} cabang unik")

        sub26 = df[df["Tahun"] == 2026]
        if len(sub26) > 0:
            tot_sav = sub26["diskon_rupiah"].sum()
            avg_disc = sub26["efisiensi_diskon_pct"].mean()
            avg_dur = sub26["durasi_nego_hari"].mean()
            print(f"\nSTATISTIK KOHORT TERBARU (TAHUN 2026 - JANUARI S/D AGUSTUS):")
            print(f"   - Total Transaksi Cabang    : {len(sub26)} cabang")
            print(f"   - Total Penghematan Sewa    : {rupiah(tot_sav)}")
            print(f"   - Rerata Diskon Sewa        : {avg_disc:.2f}%")
            print(f"   - Rerata Durasi Negosiasi   : {avg_dur:.1f} hari")
            print(f"   - Rekap Penugasan Personil  :")
            for n_name, n_cnt in sub26["negosiator_analisis"].value_counts().items():
                sub_n = sub26[sub26["negosiator_analisis"] == n_name]
                sav_n = sub_n["diskon_rupiah"].sum()
                disc_n = sub_n["efisiensi_diskon_pct"].mean()
                dur_n = sub_n["durasi_nego_hari"].mean()
                print(f"      • {n_name:<10}: {n_cnt:>3} deal | Saving: {rupiah(sav_n):<16} | Rerata Diskon: {disc_n:>5.2f}% | Durasi: {dur_n:>4.1f} hr")

        print("=" * 95 + "\n")
    return df


if __name__ == "__main__":
    run_cleansing()
