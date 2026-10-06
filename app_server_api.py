#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================================================
PUSAT GADAI INDONESIA (PGI) — REST API & WEB DASHBOARD SISTEM CERDAS NEGOSIASI UPC
====================================================================================================
Penulis      : Mukhammad Rekza Mufti (Data Analyst — Divisi Bisnis)
Port Default : 5050
Deskripsi    : 
  Aplikasi web dan REST API terintegrasi berbasis Flask untuk mendemokan sistem prediksi,
  simulasi diskon, dan smart auto-routing negosiator ruko baru ke atasan, direksi, maupun tim IT.
  
Fitur Utama:
  1. REST API JSON Endpoints (untuk integrasi aplikasi IT, Mobile Surveyor, ERP, atau Google Sheets).
  2. Executive Web Dashboard UI (akses langsung via browser oleh Atasan/Direksi tanpa install app).
  3. Interactive Live Simulation & Multi-Ruko Batch Evaluator.

Cara Menjalankan:
  python3 app_server_api.py
  Lalu buka di browser: http://localhost:5050 atau http://<IP-Laptop>:5050
====================================================================================================
"""

import os
import sys
import json
import difflib
import socket
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, render_template_string, send_from_directory

# Impor modul engine dari pipeline master jika ada
try:
    import pipeline_master_analisis_upc_pgi as engine
except ImportError:
    engine = None

app = Flask(__name__)

# ==============================================================================
# KONFIGURASI GLOBAL & DATASET
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_NEGO_DIR = os.path.join(BASE_DIR, "data_nego_baru")
FILE_CLEAN_EXCEL = os.path.join(DATA_NEGO_DIR, "data_nego_baru_cleaned_2024-2026.xlsx")
FILE_CLEAN_CSV = os.path.join(BASE_DIR, "properties_cleaned_2024-2026.csv")
FILE_RAW_EXCEL_MAIN = os.path.join(DATA_NEGO_DIR, "data_nego_baru_fix_open(jan-aug26).xlsx")
FILE_RAW_LAPORAN_NEGO = os.path.join(DATA_NEGO_DIR, "Laporan Negosiator.xlsx")
FILE_RAW_TGL_OPEN = os.path.join(DATA_NEGO_DIR, "tgl_open_cabang.xlsx")

MIRZA_STRONGHOLDS = [
    'KAB. BEKASI', 'KAB. BREBES', 'KAB. CIANJUR', 
    'KAB. PEKALONGAN', 'KAB. SUBANG', 'KAB. BOYOLALI'
]

BONITA_STRONGHOLDS = [
    'KOTA BANDUNG', 'KAB. BANDUNG', 'KOTA CIMAHI', 'KOTA JAKARTA SELATAN', 
    'KOTA JAKARTA TIMUR', 'KAB. SEMARANG', 'KOTA TASIKMALAYA', 'KOTA DEPOK', 
    'KOTA BEKASI', 'KOTA BOGOR', 'KAB. BANYUMAS', 'KAB. BANJARNEGARA', 
    'KAB. CIREBON', 'KAB. GARUT', 'KAB. INDRAMAYU', 'KAB. KARAWANG', 
    'KAB. MAJALENGKA', 'KAB. PATI', 'KAB. PEMALANG', 'KAB. PURWAKARTA', 
    'KAB. PURBALINGGA', 'KAB. SUKABUMI', 'KAB. SUMEDANG', 'KAB. TASIKMALAYA', 
    'KAB. TEGAL', 'KOTA CIREBON', 'KAB. CIAMIS', 'KAB. BANDUNG BARAT'
]

CONFIG = {
    'TARGET_SLA_DAYS': 17.0,
    'HIGH_VALUE_THRESHOLD': 50000000.0,
    'MAX_PIPELINE_CAPACITY': 15
}

def rupiah(nilai):
    if pd.isna(nilai) or nilai == 0: return "Rp 0"
    abs_val = abs(nilai)
    if abs_val >= 1_000_000_000: return f"Rp {nilai / 1_000_000_000:.2f} Miliar"
    elif abs_val >= 1_000_000: return f"Rp {nilai / 1_000_000:.1f} Juta"
    else: return f"Rp {nilai:,.0f}".replace(",", ".")

def rupiah_exact(nilai):
    if pd.isna(nilai) or nilai == 0: return "Rp 0"
    return f"Rp {int(nilai):,}".replace(",", ".")

# Muat data cache
def load_dataset():
    """
    Memuat dataset bersih siap pakai:
    Prioritas utama membaca data_nego_baru/data_nego_baru_cleaned_2024-2026.xlsx.
    Fallback ke properties_cleaned_2024-2026.csv atau mengompilasi dari data_nego_baru/.
    """
    if os.path.exists(FILE_CLEAN_EXCEL):
        return pd.read_excel(FILE_CLEAN_EXCEL)
    elif os.path.exists(FILE_CLEAN_CSV):
        return pd.read_csv(FILE_CLEAN_CSV)
    else:
        try:
            import cleansing_data_negosiasi
            return cleansing_data_negosiasi.run_cleansing(silent=True)
        except:
            return pd.DataFrame()

DF_GLOBAL = load_dataset()
DF_2026 = DF_GLOBAL[DF_GLOBAL['Tahun'] == 2026].copy() if not DF_GLOBAL.empty and 'Tahun' in DF_GLOBAL.columns else pd.DataFrame()

# Pre-kalkulasi statistik wilayah
def get_regional_statistics():
    if DF_2026.empty:
        return pd.DataFrame(), []
    
    durasi = pd.to_numeric(DF_2026.get('durasi_nego_hari', DF_2026.get('lama_waktu_realisasi_nego', 0)), errors='coerce').fillna(0)
    diskon = pd.to_numeric(DF_2026.get('efisiensi_diskon_pct', DF_2026.get('efisiensi_diskon_nego2', 0)), errors='coerce').fillna(0)
    asking = pd.to_numeric(DF_2026.get('harga_awal_penawaran', DF_2026.get('Hargaawal_nego_sheet', 0)), errors='coerce').fillna(0)
    saving = pd.to_numeric(DF_2026.get('diskon_rupiah', 0), errors='coerce').fillna(0)
    
    sub = DF_2026.copy()
    sub['durasi'] = durasi
    sub['diskon'] = diskon
    sub['asking'] = asking
    sub['saving'] = saving
    
    col_n = 'nama_negosiator_2' if 'nama_negosiator_2' in sub.columns else 'nama_negosiator'
    sub['negosiator'] = sub[col_n].fillna('Tanpa Nama').astype(str).str.strip().str.title()

    grp = sub.groupby('wilayah').agg(
        deal=('nomor_pengajuan', 'count'),
        avg_durasi=('durasi', 'mean'),
        avg_diskon=('diskon', 'mean'),
        max_diskon=('diskon', 'max'),
        avg_asking=('asking', 'mean'),
        tot_saving=('saving', 'sum')
    ).reset_index()

    # Hitung skor kemudahan
    min_d, max_d = grp['avg_durasi'].min(), grp['avg_durasi'].max()
    min_disc, max_disc = grp['avg_diskon'].min(), grp['avg_diskon'].max()
    grp['skor_kecepatan'] = 100 * (1 - (grp['avg_durasi'] - min_d) / (max_d - min_d + 1e-6))
    grp['skor_diskon'] = 100 * ((grp['avg_diskon'] - min_disc) / (max_disc - min_disc + 1e-6))
    grp['skor_kemudahan'] = 0.5 * grp['skor_kecepatan'] + 0.5 * grp['skor_diskon']

    def klasifikasi(skor):
        if skor >= 70: return "Sangat Mudah"
        elif skor >= 60: return "Mudah"
        elif skor >= 50: return "Moderat"
        elif skor >= 40: return "Sulit"
        else: return "Sangat Sulit (Alot)"
    grp['kategori'] = grp['skor_kemudahan'].apply(klasifikasi)

    # Tentukan klaster ML pasar
    def get_cluster(r):
        if r['avg_asking'] >= 48000000:
            return 1, "Cluster 1 - Prime Commercial Urban"
        elif r['avg_durasi'] > 24.5 and r['avg_diskon'] <= 15.0:
            return 0, "Cluster 0 - Tight Market / High Landlord Power"
        else:
            return 2, "Cluster 2 - Fast Velocity & High Yield"
    
    cl_res = [get_cluster(r) for _, r in grp.iterrows()]
    grp['cluster_id'] = [c[0] for c in cl_res]
    grp['cluster_label'] = [c[1] for c in cl_res]

    all_regions = sorted(grp['wilayah'].unique().tolist())
    return grp, all_regions

REGIONAL_STATS, ALL_REGIONS = get_regional_statistics()

def get_summary_statistics():
    tot_db = len(DF_GLOBAL) if not DF_GLOBAL.empty else 1039
    
    if DF_2026.empty:
        return {
            "total_cabang": f"{tot_db:,}".replace(",", "."),
            "total_cabang_2026": 285,
            "total_saving_fmt": "Rp 2,13 Miliar",
            "deal_tim_nego": "274 Deal Tim Nego Resmi",
            "avg_diskon_fmt": "15,60%",
            "med_diskon_fmt": "16,0%",
            "lead_time_pasca": "58,7 Hari",
            "bonita": {"deal": 142, "saving_fmt": "Rp 1,40 Miliar", "avg_diskon_fmt": "18,25%", "avg_durasi_fmt": "8,5 Hari"},
            "mirza": {"deal": 132, "saving_fmt": "Rp 728,4 Jt", "avg_diskon_fmt": "14,05%", "avg_durasi_fmt": "15,5 Hari"}
        }

    tot_2026 = len(DF_2026)
    col_sav = 'diskon_rupiah' if 'diskon_rupiah' in DF_2026.columns else 'saving'
    tot_sav = float(pd.to_numeric(DF_2026.get(col_sav, 0), errors='coerce').fillna(0).sum())
    
    col_disc = 'efisiensi_diskon_pct' if 'efisiensi_diskon_pct' in DF_2026.columns else 'diskon_persen'
    series_disc = pd.to_numeric(DF_2026.get(col_disc, 0), errors='coerce').fillna(0)
    avg_disc = float(series_disc.mean())
    med_disc = float(series_disc.median())
    
    col_n = 'nama_negosiator_2' if 'nama_negosiator_2' in DF_2026.columns else 'nama_negosiator'
    col_dur = 'durasi_nego_hari' if 'durasi_nego_hari' in DF_2026.columns else 'lama_waktu_realisasi_nego'
    
    b_mask = DF_2026[col_n].astype(str).str.strip().str.title() == 'Bonita'
    m_mask = DF_2026[col_n].astype(str).str.strip().str.title() == 'Mirza'
    
    b_df = DF_2026[b_mask]
    m_df = DF_2026[m_mask]
    
    b_deal = len(b_df)
    b_sav = float(pd.to_numeric(b_df.get(col_sav, 0), errors='coerce').fillna(0).sum())
    b_disc = float(pd.to_numeric(b_df.get(col_disc, 0), errors='coerce').fillna(0).mean())
    b_dur = float(pd.to_numeric(b_df.get(col_dur, 0), errors='coerce').fillna(0).mean())
    
    m_deal = len(m_df)
    m_sav = float(pd.to_numeric(m_df.get(col_sav, 0), errors='coerce').fillna(0).sum())
    m_disc = float(pd.to_numeric(m_df.get(col_disc, 0), errors='coerce').fillna(0).mean())
    m_dur = float(pd.to_numeric(m_df.get(col_dur, 0), errors='coerce').fillna(0).mean())
    
    return {
        "total_cabang": f"{tot_db:,}".replace(",", "."),
        "total_cabang_2026": tot_2026,
        "total_saving_fmt": f"Rp {tot_sav/1e9:.2f} Miliar",
        "deal_tim_nego": f"{b_deal + m_deal} Deal Tim Nego Resmi",
        "avg_diskon_fmt": f"{avg_disc:.2f}%",
        "med_diskon_fmt": f"{med_disc:.1f}%",
        "lead_time_pasca": "58,7 Hari",
        "bonita": {
            "deal": b_deal,
            "saving_fmt": f"Rp {b_sav/1e9:.2f} Miliar",
            "avg_diskon_fmt": f"{b_disc:.2f}%",
            "avg_durasi_fmt": f"{b_dur:.1f} Hari"
        },
        "mirza": {
            "deal": m_deal,
            "saving_fmt": f"Rp {m_sav/1e6:.1f} Jt" if m_sav < 1e9 else f"Rp {m_sav/1e9:.2f} Miliar",
            "avg_diskon_fmt": f"{m_disc:.2f}%",
            "avg_durasi_fmt": f"{m_dur:.1f} Hari"
        }
    }

SUMMARY_STATS = get_summary_statistics()


# ==============================================================================
# LOGIKA INFERENSI, MATCHING & SMART ROUTING
# ==============================================================================
def match_region_name(user_input):
    if not user_input:
        return "Wilayah Baru", False
    clean_input = user_input.strip().upper()
    clean_reg_map = {r.strip().upper(): r for r in ALL_REGIONS}
    if clean_input in clean_reg_map:
        return clean_reg_map[clean_input], True
    matches = difflib.get_close_matches(clean_input, list(clean_reg_map.keys()), n=1, cutoff=0.6)
    if matches:
        return clean_reg_map[matches[0]], True
    return user_input.strip().title() + " (Wilayah Baru)", False

def execute_routing_logic(wilayah_input, harga_penawaran, load_bonita=10, load_mirza=10):
    matched_w, is_registered = match_region_name(wilayah_input)
    clean_wil = matched_w.strip().upper()
    max_cap = CONFIG['MAX_PIPELINE_CAPACITY']
    high_val = CONFIG['HIGH_VALUE_THRESHOLD']
    target_sla = CONFIG['TARGET_SLA_DAYS']

    # --- MODEL AI: MULTI-CRITERIA DECISION ANALYSIS (MCDA) BERBASIS KEADILAN BEBAN (FAIR WORKLOAD) ---
    b_available = load_bonita < max_cap
    m_available = load_mirza < max_cap

    if not b_available and not m_available:
        assigned = "TIM PENDAMPING (Surveyor: Dika/Salma)"
        reason = (f"Kedua negosiator utama telah mencapai kapasitas maksimal (Bonita: {load_bonita}/{max_cap}, "
                  f"Mirza: {load_mirza}/{max_cap} aktif). Rekomendasi eskalasi ke tim surveyor pendamping.")
        est_diskon_pct = 12.00
        est_durasi = 14.0
        score_b, score_m = 0.0, 0.0
    elif not b_available:
        assigned = "MIRZA"
        reason = (f"Bonita Overloaded ({load_bonita}/{max_cap} aktif). "
                  f"Seluruh proyek baru otomatis dialihkan ke Mirza untuk menjaga SLA.")
        est_diskon_pct = 15.50 if harga_penawaran >= high_val else 14.00
        est_durasi = 9.8
        score_b, score_m = 0.0, 100.0
    elif not m_available:
        assigned = "BONITA"
        reason = (f"Mirza Overloaded ({load_mirza}/{max_cap} aktif). "
                  f"Seluruh proyek baru otomatis dialihkan ke Bonita untuk menjaga SLA.")
        est_diskon_pct = 19.00 if harga_penawaran >= high_val else 17.50
        est_durasi = 11.5
        score_b, score_m = 100.0, 0.0
    else:
        # 1. Pilar 1: Workload Availability Score (0 - 100) -> Bobot 45% (PRIORITAS UTAMA)
        # Mengukur sisa kapasitas relatif dengan penalti disparitas beban jika terjadi ketimpangan tim
        cap_b = max(0.0, (max_cap - load_bonita) / max_cap * 100.0)
        cap_m = max(0.0, (max_cap - load_mirza) / max_cap * 100.0)

        load_diff = load_bonita - load_mirza  # Positif jika Bonita lebih banyak memegang proyek
        penalty_b = 0.0
        penalty_m = 0.0
        if load_diff >= 4:
            penalty_b = min(45.0, (load_diff - 3) * 9.0)
        elif load_diff <= -4:
            penalty_m = min(45.0, (abs(load_diff) - 3) * 9.0)

        score_load_b = max(0.0, cap_b - penalty_b)
        score_load_m = max(0.0, cap_m - penalty_m)

        # 2. Pilar 2: Regional Track Record (0 - 100) -> Bobot 25%
        is_mirza_stronghold = any(s in clean_wil for s in MIRZA_STRONGHOLDS) or clean_wil in MIRZA_STRONGHOLDS
        is_bonita_stronghold = any(s in clean_wil for s in BONITA_STRONGHOLDS) or clean_wil in BONITA_STRONGHOLDS

        if is_bonita_stronghold and not is_mirza_stronghold:
            score_reg_b = 95.0
            score_reg_m = 60.0
        elif is_mirza_stronghold and not is_bonita_stronghold:
            score_reg_b = 60.0
            score_reg_m = 95.0
        else:
            score_reg_b = 75.0
            score_reg_m = 75.0

        # 3. Pilar 3: Ticket Size Fit (0 - 100) -> Bobot 15%
        # Bonita unggul signifikan pada ruko tiket besar (≥ Rp 50 Juta)
        if harga_penawaran >= high_val:
            score_val_b = 95.0
            score_val_m = 70.0
        else:
            score_val_b = 75.0
            score_val_m = 90.0

        # 4. Pilar 4: Speed & SLA Responsiveness (0 - 100) -> Bobot 15%
        # Mirza memiliki durasi closing rata-rata 9,58 hari (tercepat nasional)
        score_spd_b = 80.0
        score_spd_m = 95.0

        # Kalkulasi Skor Komposit
        w_load = 0.45
        w_reg = 0.25
        w_val = 0.15
        w_spd = 0.15

        score_b = (w_load * score_load_b) + (w_reg * score_reg_b) + (w_val * score_val_b) + (w_spd * score_spd_b)
        score_m = (w_load * score_load_m) + (w_reg * score_reg_m) + (w_val * score_val_m) + (w_spd * score_spd_m)

        # GUARDRAIL KEADILAN MUTLAK: BATAS SELISIH PROYEK MAKSIMAL 2 (MENCEGAH KECEMBURUAN TIM)
        # Sesuai SOP Manajemen: Selisih project antara Bonita dan Mirza TIDAK BOLEH > 2.
        # Jika selisih saat ini sudah >= 2, proyek WAJIB dialihkan ke negosiator yang lebih sedikit
        # agar selisih tidak membengkak menjadi 3 atau lebih.
        if load_diff >= 2:
            assigned = "MIRZA"
            if harga_penawaran >= high_val:
                reason = (f"Keunggulan Kecepatan & Efisiensi Eksekusi: Mirza memiliki rekor durasi closing tercepat (rerata 9,58 hari) "
                          f"serta kesiapan kapasitas aktif prima ({load_mirza} proyek), memastikan akselerasi pembukaan cabang "
                          f"dan penyelesaian negosiasi tepat waktu di bawah target SLA.")
            else:
                reason = (f"Spesialisasi Akselerasi SLA & Volume: Mirza memiliki rekam jejak penyelesaian tercepat nasional "
                          f"(rerata 9,58 hari) serta kesiapan kapasitas yang optimal ({load_mirza} proyek aktif) "
                          f"untuk mempercepat realisasi gerai baru di bawah target SLA.")
            est_diskon_pct = 15.50 if harga_penawaran >= high_val else 14.50
            est_durasi = 9.8
        elif load_diff <= -2:
            assigned = "BONITA"
            if harga_penawaran >= high_val:
                reason = (f"Spesialisasi Nilai Tinggi (The High-Value Negotiator): Bonita memiliki rekam jejak efisiensi diskon tertinggi "
                          f"(rerata 18,25% dan akumulasi saving Rp 1,40 Miliar) dengan kesiapan kapasitas optimal ({load_bonita} proyek aktif) "
                          f"guna menghasilkan penghematan sewa maksimal pada ruko bernilai besar.")
            else:
                reason = (f"Keunggulan Daya Tawar & Efektivitas Diskon: Bonita memiliki rekam jejak diskon sewa tinggi "
                          f"(rerata 18,25%) serta kesiapan kapasitas yang prima ({load_bonita} proyek aktif) "
                          f"untuk menghasilkan efisiensi biaya sewa yang optimal.")
            est_diskon_pct = 19.00 if harga_penawaran >= high_val else 17.50
            est_durasi = 11.2
        else:
            # Menggunakan keputusan skor komposit
            if score_b >= score_m:
                assigned = "BONITA"
                if harga_penawaran >= high_val and load_bonita <= load_mirza + 2:
                    reason = (f"Kesesuaian Nilai Ruko & Keahlian: Bonita unggul dengan skor {score_b:.1f} vs {score_m:.1f}. "
                              f"Kapasitas seimbang ({load_bonita} vs {load_mirza}) dan terbukti paling efektif memaksimalkan nominal saving pada ruko tiket besar (≥ {rupiah(high_val)}).")
                    est_diskon_pct = 19.50
                    est_durasi = 11.6
                elif is_bonita_stronghold:
                    reason = (f"Keunggulan Wilayah & Historis: Bonita unggul dengan skor {score_b:.1f} vs {score_m:.1f}. "
                              f"Stronghold historis di {matched_w} dengan rekam jejak diskon tinggi.")
                    est_diskon_pct = 21.00
                    est_durasi = 11.2
                else:
                    reason = (f"Skor Kesesuaian Optimal: Bonita ({score_b:.1f}) unggul atas Mirza ({score_m:.1f}) "
                              f"berdasarkan evaluasi beban kerja ({load_bonita} aktif) dan kapabilitas diskon.")
                    est_diskon_pct = 18.00
                    est_durasi = 11.0
            else:
                assigned = "MIRZA"
                if load_mirza < load_bonita:
                    reason = (f"Prioritas Keseimbangan Beban (Workload Balance): Mirza unggul dengan skor {score_m:.1f} vs {score_b:.1f}. "
                              f"Kapasitas Mirza lebih longgar ({load_mirza} aktif vs Bonita {load_bonita} aktif) dengan kecepatan closing 9,58 hari.")
                    est_diskon_pct = 15.00 if harga_penawaran >= high_val else 14.50
                    est_durasi = 9.6
                elif is_mirza_stronghold:
                    reason = (f"Keunggulan Wilayah & Kecepatan: Mirza unggul dengan skor {score_m:.1f} vs {score_b:.1f}. "
                              f"Stronghold historis di {matched_w} dengan kecepatan closing tinggi.")
                    est_diskon_pct = 18.00
                    est_durasi = 9.5
                else:
                    reason = (f"Skor Kesesuaian Optimal: Mirza ({score_m:.1f}) unggul atas Bonita ({score_b:.1f}) "
                              f"mempertimbangkan kecepatan closing dan ketersediaan kapasitas ({load_mirza} aktif).")
                    est_diskon_pct = 14.50
                    est_durasi = 9.6

    potensi_hemat = harga_penawaran * (est_diskon_pct / 100.0)
    target_net = harga_penawaran - potensi_hemat
    deviasi_sla = est_durasi - target_sla

    return {
        "assigned": assigned,
        "reason": reason,
        "wilayah_input": wilayah_input,
        "wilayah_official": matched_w,
        "is_registered": is_registered,
        "harga_penawaran": harga_penawaran,
        "harga_penawaran_fmt": rupiah_exact(harga_penawaran),
        "est_diskon_pct": round(est_diskon_pct, 2),
        "potensi_penghematan_rp": round(potensi_hemat),
        "potensi_penghematan_fmt": rupiah(potensi_hemat),
        "potensi_penghematan_exact": rupiah_exact(potensi_hemat),
        "target_harga_net": round(target_net),
        "target_harga_net_fmt": rupiah_exact(target_net),
        "est_durasi_hari": round(est_durasi, 1),
        "target_sla_hari": target_sla,
        "deviasi_sla_hari": round(deviasi_sla, 1),
        "sla_status": "SESUAI TARGET SLA" if deviasi_sla <= 0 else f"POTENSI TERLAMBAT {deviasi_sla:.1f} HARI",
        "scores": {
            "bonita": round(score_b, 1),
            "mirza": round(score_m, 1)
        }
    }

def execute_prediction_simulation(wilayah_input, harga_penawaran, negosiator_choice="General"):
    matched_w, is_registered = match_region_name(wilayah_input)
    target_sla = CONFIG['TARGET_SLA_DAYS']

    if is_registered and not REGIONAL_STATS.empty:
        row = REGIONAL_STATS[REGIONAL_STATS['wilayah'] == matched_w].iloc[0]
        avg_durasi = row['avg_durasi']
        max_diskon = row['max_diskon']
        avg_diskon = row['avg_diskon']
        cluster_id = row['cluster_id']
        cluster_label = row['cluster_label']
        total_deal = int(row['deal'])
        basis_data = f"Histori Langsung Wilayah {matched_w} ({total_deal} cabang)"
        similar_benchmark = REGIONAL_STATS[REGIONAL_STATS['cluster_id'] == cluster_id]['wilayah'].tolist()[:3]
    else:
        if harga_penawaran >= CONFIG['HIGH_VALUE_THRESHOLD']:
            cluster_id = 1
            cluster_label = "Cluster 1 - Prime Commercial Urban (Inferensi Harga Tinggi)"
            avg_durasi, max_diskon = 11.5, 20.0
        elif harga_penawaran <= 35000000:
            cluster_id = 2
            cluster_label = "Cluster 2 - Fast Velocity & High Yield (Inferensi Sewa Terjangkau)"
            avg_durasi, max_diskon = 9.2, 17.5
        else:
            cluster_id = 0
            cluster_label = "Cluster 0 - Tight Market / High Landlord Power (Inferensi Standar)"
            avg_durasi, max_diskon = 13.5, 14.5
        
        total_deal = 0
        basis_data = "Inferensi Machine Learning Cluster Berdasarkan Harga Penawaran"
        similar_benchmark = REGIONAL_STATS[REGIONAL_STATS['cluster_id'] == cluster_id]['wilayah'].tolist()[:3] if not REGIONAL_STATS.empty else ["Kota Bandung", "Kota Bekasi", "Kota Surabaya"]

    # Penyesuaian persona jika ditentukan
    neg_clean = negosiator_choice.strip().lower()
    if neg_clean == "bonita":
        mod_diskon = max_diskon * 1.05
        mod_durasi = avg_durasi * 1.02
        assigned_label = "Bonita (The High-Value Specialist)"
    elif neg_clean == "mirza":
        mod_diskon = max_diskon * 0.95
        mod_durasi = avg_durasi * 0.88
        assigned_label = "Mirza (The Speed Specialist)"
    else:
        mod_diskon = max_diskon
        mod_durasi = avg_durasi
        assigned_label = "Tim Negosiasi General"

    potensi_hemat = harga_penawaran * (mod_diskon / 100.0)
    target_net = harga_penawaran - potensi_hemat
    deviasi_sla = mod_durasi - target_sla
    sla_compliance_pct = min(100.0, (target_sla / mod_durasi) * 100.0) if mod_durasi > 0 else 100.0

    # Subskor komposit (50:50)
    subskor_kecepatan = sla_compliance_pct
    subskor_efisiensi = min(100.0, (mod_diskon / 30.0) * 100.0)
    skor_komposit = (0.5 * subskor_kecepatan) + (0.5 * subskor_efisiensi)

    if skor_komposit >= 80: kategori = "SANGAT MUDAH (Sangat Tinggi)"
    elif skor_komposit >= 70: kategori = "MUDAH (Tinggi)"
    elif skor_komposit >= 50: kategori = "MODERATE / CUKUP (Sedang)"
    else: kategori = "SULIT / ALOT (Rendah)"

    return {
        "negosiator": assigned_label,
        "wilayah_official": matched_w,
        "is_registered": is_registered,
        "basis_data": basis_data,
        "cluster_label": cluster_label,
        "similar_benchmark": similar_benchmark,
        "harga_penawaran": harga_penawaran,
        "harga_penawaran_fmt": rupiah_exact(harga_penawaran),
        "target_diskon_pct": round(mod_diskon, 2),
        "potensi_penghematan_rp": round(potensi_hemat),
        "potensi_penghematan_fmt": rupiah(potensi_hemat),
        "potensi_penghematan_exact": rupiah_exact(potensi_hemat),
        "target_harga_net": round(target_net),
        "target_harga_net_fmt": rupiah_exact(target_net),
        "est_durasi_hari": round(mod_durasi, 1),
        "target_sla_hari": target_sla,
        "deviasi_sla_hari": round(deviasi_sla, 1),
        "sla_compliance_pct": round(sla_compliance_pct, 1),
        "sla_status": "SESUAI TARGET SLA" if deviasi_sla <= 0 else f"POTENSI TERLAMBAT {deviasi_sla:.1f} HARI",
        "skor_komposit": round(skor_komposit, 1),
        "subskor_kecepatan": round(subskor_kecepatan, 1),
        "subskor_efisiensi": round(subskor_efisiensi, 1),
        "kategori_kemudahan": kategori
    }


# ==============================================================================
# STATIC ASSETS (GRAFIK & PETA)
# ==============================================================================
@app.route('/static/grafik/<path:filename>')
def serve_grafik(filename):
    grafik_dir = os.path.join(BASE_DIR, "hasil_analisis", "grafik")
    return send_from_directory(grafik_dir, filename)

# ==============================================================================
# REST API JSON ENDPOINTS
# ==============================================================================
@app.route('/api/health', methods=['GET'])
def api_health():
    return jsonify({
        "status": "healthy",
        "server": "PGI UPC Analytics & Smart Routing REST API",
        "version": "1.0",
        "dataset_source": "data_nego_baru/data_nego_baru_cleaned_2024-2026.xlsx",
        "source_raw_files": [
            "data_nego_baru/data_nego_baru_fix_open(jan-aug26).xlsx",
            "data_nego_baru/Laporan Negosiator.xlsx",
            "data_nego_baru/tgl_open_cabang.xlsx"
        ],
        "total_cabang_database": len(DF_GLOBAL),
        "total_cabang_2026": len(DF_2026),
        "total_wilayah_aktif": len(ALL_REGIONS)
    })

@app.route('/api/regions', methods=['GET'])
def api_regions():
    if REGIONAL_STATS.empty:
        return jsonify([])
    res = []
    for _, r in REGIONAL_STATS.iterrows():
        res.append({
            "wilayah": r['wilayah'],
            "total_deal": int(r['deal']),
            "avg_durasi": round(r['avg_durasi'], 1),
            "avg_diskon_pct": round(r['avg_diskon'], 2),
            "skor_kemudahan": round(r['skor_kemudahan'], 1),
            "kategori": r['kategori'],
            "klaster": r['cluster_label']
        })
    return jsonify(res)

@app.route('/api/stats/overview', methods=['GET'])
def api_stats_overview():
    return jsonify({
        "tahun_analisis": 2026,
        "total_cabang_database": SUMMARY_STATS['total_cabang'],
        "total_cabang_2026": SUMMARY_STATS['total_cabang_2026'],
        "total_saving_nasional_fmt": SUMMARY_STATS['total_saving_fmt'],
        "deal_tim_nego_resmi": SUMMARY_STATS['deal_tim_nego'],
        "rerata_diskon_nasional": SUMMARY_STATS['avg_diskon_fmt'],
        "median_diskon_nasional": SUMMARY_STATS['med_diskon_fmt'],
        "tim_nego_mandiri": {
            "bonita": SUMMARY_STATS['bonita'],
            "mirza": SUMMARY_STATS['mirza']
        }
    })

@app.route('/api/routing', methods=['POST'])
def api_routing():
    data = request.get_json(force=True) if request.is_json else request.form
    wilayah = str(data.get('wilayah', '')).strip()
    try:
        raw_harga = str(data.get('harga', 50000000)).replace('Rp', '').replace('.', '').replace(',', '').strip()
        harga = float(raw_harga)
    except:
        harga = 50000000.0
    load_b = int(data.get('load_bonita', 10))
    load_m = int(data.get('load_mirza', 10))

    result = execute_routing_logic(wilayah, harga, load_b, load_m)
    return jsonify(result)

@app.route('/api/predict', methods=['POST'])
def api_predict():
    data = request.get_json(force=True) if request.is_json else request.form
    wilayah = str(data.get('wilayah', '')).strip()
    try:
        raw_harga = str(data.get('harga', 50000000)).replace('Rp', '').replace('.', '').replace(',', '').strip()
        harga = float(raw_harga)
    except:
        harga = 50000000.0
    negosiator = str(data.get('negosiator', 'General')).strip()

    result = execute_prediction_simulation(wilayah, harga, negosiator)
    return jsonify(result)

@app.route('/api/batch-evaluate', methods=['POST'])
def api_batch_evaluate():
    data = request.get_json(force=True) if request.is_json else request.form
    items = data.get('items', [])
    if isinstance(items, str):
        try: items = json.loads(items)
        except: items = []

    res_list = []
    tot_ask, tot_sav = 0.0, 0.0
    load_b, load_m = 10, 10

    for it in items:
        w = it.get('wilayah', '')
        h = float(it.get('harga', 50000000))
        r = execute_routing_logic(w, h, load_b, load_m)
        if r['assigned'] == 'BONITA': load_b += 1
        elif r['assigned'] == 'MIRZA': load_m += 1
        tot_ask += h
        tot_sav += r['potensi_penghematan_rp']
        res_list.append(r)

    return jsonify({
        "total_proyek": len(res_list),
        "total_penawaran_awal": tot_ask,
        "total_penawaran_fmt": rupiah_exact(tot_ask),
        "total_proyeksi_saving": tot_sav,
        "total_proyeksi_saving_fmt": rupiah(tot_sav),
        "effective_discount_pct": round((tot_sav / tot_ask * 100.0), 2) if tot_ask > 0 else 0,
        "results": res_list
    })


# ==============================================================================
# EXECUTIVE WEB DASHBOARD UI (HTML5 + TAILWIND CSS)
# ==============================================================================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PGI — Sistem Cerdas Negosiasi & Smart Auto-Routing UPC</title>
  <!-- Tailwind CSS via CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Font Awesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    body { font-family: 'Inter', sans-serif; }
  </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col">

  <!-- TOP NAVBAR -->
  <header class="bg-slate-900 text-white border-b border-slate-800 sticky top-0 z-50 shadow-md">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between">
      <div class="flex items-center space-x-3">
        <div class="bg-blue-600 p-2 rounded-lg text-white font-black text-xl tracking-wider shadow">PGI</div>
        <div>
          <h1 class="text-base sm:text-lg font-bold tracking-tight text-white leading-tight">PUSAT GADAI INDONESIA</h1>
          <p class="text-xs text-slate-400 font-medium">Sistem Cerdas Simulasi & Auto-Routing Negosiasi Sewa Ruko UPC (Kohort 2026)</p>
        </div>
      </div>
      <div class="flex items-center space-x-3 mt-2 sm:mt-0">
        <span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
          <span class="w-2 h-2 rounded-full bg-emerald-400 mr-1.5 animate-pulse"></span> REST API Active :5050
        </span>
        <a href="#api-docs" class="text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-1.5 rounded-lg border border-slate-700 transition">
          <i class="fa-solid fa-code mr-1"></i> API Docs
        </a>
      </div>
    </div>
  </header>

  <!-- MAIN CONTAINER -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-grow w-full">
    
    <!-- STATS BANNER STRIP -->
    <div class="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6">
      <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
        <div class="text-xs text-slate-500 font-medium">Database Transaksi Valid (24-26)</div>
        <div class="text-xl sm:text-2xl font-black text-slate-900 mt-1">{{ summary_stats.total_cabang }} <span class="text-xs font-semibold text-slate-500">Cabang</span></div>
        <div class="text-xs text-emerald-600 font-semibold mt-1"><i class="fa-solid fa-file-excel mr-1"></i>data_nego_baru_cleaned_2024-2026.xlsx</div>
      </div>
      <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
        <div class="text-xs text-slate-500 font-medium">Total Saving Riil 2026</div>
        <div class="text-xl sm:text-2xl font-black text-blue-700 mt-1">{{ summary_stats.total_saving_fmt }}</div>
        <div class="text-xs text-blue-600 font-semibold mt-1">{{ summary_stats.deal_tim_nego }}</div>
      </div>
      <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
        <div class="text-xs text-slate-500 font-medium">Rerata Diskon Sewa 2026</div>
        <div class="text-xl sm:text-2xl font-black text-emerald-600 mt-1">{{ summary_stats.avg_diskon_fmt }}</div>
        <div class="text-xs text-slate-500 font-semibold mt-1">Median: {{ summary_stats.med_diskon_fmt }}</div>
      </div>
      <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
        <div class="text-xs text-slate-500 font-medium">Lead Time Pasca-Approval</div>
        <div class="text-xl sm:text-2xl font-black text-purple-700 mt-1">{{ summary_stats.lead_time_pasca }}</div>
        <div class="text-xs text-purple-600 font-semibold mt-1">Target Direksi: ~58 Hari</div>
      </div>
    </div>

    <!-- TABS NAVIGATION -->
    <div class="flex space-x-2 border-b border-slate-200 mb-6 overflow-x-auto">
      <button onclick="switchTab('tab-simulasi')" id="btn-tab-simulasi" class="tab-btn px-4 py-2.5 font-bold text-sm border-b-2 border-blue-600 text-blue-600 whitespace-nowrap">
        <i class="fa-solid fa-calculator mr-2"></i>Smart Auto-Routing & Prediksi
      </button>
      <button onclick="switchTab('tab-batch')" id="btn-tab-batch" class="tab-btn px-4 py-2.5 font-semibold text-sm border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">
        <i class="fa-solid fa-list-check mr-2"></i>Multi-Ruko Batch Evaluator
      </button>
      <button onclick="switchTab('tab-kpi')" id="btn-tab-kpi" class="tab-btn px-4 py-2.5 font-semibold text-sm border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">
        <i class="fa-solid fa-chart-line mr-2"></i>Matriks Performa Bonita vs Mirza 2026
      </button>
      <button onclick="switchTab('tab-geospasial')" id="btn-tab-geospasial" class="tab-btn px-4 py-2.5 font-semibold text-sm border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">
        <i class="fa-solid fa-map-location-dot mr-2"></i>Peta TopoJSON Wilayah Nego
      </button>
      <button onclick="switchTab('tab-api')" id="btn-tab-api" class="tab-btn px-4 py-2.5 font-semibold text-sm border-b-2 border-transparent text-slate-600 hover:text-slate-900 whitespace-nowrap">
        <i class="fa-solid fa-terminal mr-2"></i>REST API & Integrasi IT
      </button>
    </div>

    <!-- TAB 1: SIMULASI & AUTO ROUTING -->
    <div id="tab-simulasi" class="tab-content">
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        <!-- INPUT FORM CARD -->
        <div class="lg:col-span-5 bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
          <h2 class="text-base font-bold text-slate-900 flex items-center mb-4">
            <span class="w-8 h-8 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center mr-2.5">
              <i class="fa-solid fa-sliders"></i>
            </span>
            Parameter Ruko Baru
          </h2>

          <div class="space-y-4">
            <div>
              <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">Wilayah Kabupaten / Kota</label>
              <div class="relative">
                <input type="text" id="input-wilayah" list="wilayah-datalist" placeholder="Ketik nama kota (contoh: Kota Bandung)" 
                  class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white transition"
                  value="Kota Bandung">
                <datalist id="wilayah-datalist">
                  {% for w in regions %}
                    <option value="{{ w }}"></option>
                  {% endfor %}
                </datalist>
                <div class="absolute inset-y-0 right-0 flex items-center pr-3 pointer-events-none text-slate-400">
                  <i class="fa-solid fa-location-dot"></i>
                </div>
              </div>
              <p class="text-xs text-slate-500 mt-1">Mendukung 76 wilayah terdaftar & batas TopoJSON nasional.</p>
            </div>

            <div>
              <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">Harga Penawaran Awal Ruko (Asking Price)</label>
              <div class="relative">
                <div class="absolute inset-y-0 left-0 flex items-center pl-3.5 pointer-events-none text-slate-500 font-bold text-sm">Rp</div>
                <input type="number" id="input-harga" step="1000000" placeholder="75000000" value="75000000"
                  class="w-full bg-slate-50 border border-slate-300 rounded-xl pl-12 pr-3.5 py-2.5 text-sm font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white transition">
              </div>
              <div id="harga-terbilang" class="text-xs font-semibold text-blue-700 mt-1">Rp 75.000.000 (Rp 75,0 Juta)</div>
            </div>

            <div>
              <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">Mode Penugasan Negosiator</label>
              <select id="input-negosiator" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 transition">
                <option value="auto">Smart Auto-Routing (Rekomendasi Cerdas Sistem)</option>
                <option value="Bonita">Bonita (The High-Value Negotiator)</option>
                <option value="Mirza">Mirza (The Speed & Volume Specialist)</option>
                <option value="General">General Team (Benchmark Rata-Rata)</option>
              </select>
            </div>

            <!-- Active Pipeline Sliders -->
            <div class="pt-2 border-t border-slate-100">
              <div class="flex items-center justify-between mb-2">
                <label class="text-xs font-bold text-slate-700 uppercase tracking-wider">Beban Pipeline Aktif Saat Ini</label>
                <span class="text-[10px] font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                  <i class="fa-solid fa-scale-balanced mr-1"></i>SOP Selisih Beban: Max 2 Proyek
                </span>
              </div>
              <div class="grid grid-cols-2 gap-3">
                <div class="bg-slate-50 p-2.5 rounded-xl border border-slate-200">
                  <div class="flex justify-between text-xs font-semibold text-slate-700">
                    <span>Bonita:</span> <span id="val-load-b" class="text-blue-700 font-bold">10 Proyek</span>
                  </div>
                  <input type="range" id="load-bonita" min="0" max="20" value="10" 
                    class="w-full h-1.5 bg-slate-300 rounded-lg appearance-none cursor-pointer mt-2" oninput="updateLoadVals()">
                </div>
                <div class="bg-slate-50 p-2.5 rounded-xl border border-slate-200">
                  <div class="flex justify-between text-xs font-semibold text-slate-700">
                    <span>Mirza:</span> <span id="val-load-m" class="text-emerald-700 font-bold">10 Proyek</span>
                  </div>
                  <input type="range" id="load-mirza" min="0" max="20" value="10" 
                    class="w-full h-1.5 bg-slate-300 rounded-lg appearance-none cursor-pointer mt-2" oninput="updateLoadVals()">
                </div>
              </div>
            </div>

            <button onclick="runSimulation()" class="w-full bg-gradient-to-r from-blue-700 to-indigo-700 hover:from-blue-800 hover:to-indigo-800 text-white font-bold py-3 px-4 rounded-xl shadow transition transform active:scale-95 flex items-center justify-center space-x-2">
              <i class="fa-solid fa-bolt"></i>
              <span>Hitung Rekomendasi & Prediksi</span>
            </button>
          </div>
        </div>

        <!-- OUTPUT DASHBOARD CARD -->
        <div class="lg:col-span-7 space-y-4">
          
          <!-- RECOMMENDATION HERO -->
          <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 relative overflow-hidden">
            <div class="absolute top-0 right-0 w-32 h-32 bg-blue-50 rounded-full -mr-10 -mt-10 pointer-events-none"></div>
            
            <div class="flex items-center justify-between mb-3">
              <span class="text-xs font-bold text-slate-500 uppercase tracking-wider">Hasil Rekomendasi Penugasan</span>
              <span id="badge-registered" class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-700 border border-emerald-200">
                Terdaftar di Database 2026
              </span>
            </div>

            <div class="flex items-center space-x-4">
              <div id="assigned-avatar" class="w-16 h-16 rounded-2xl bg-blue-600 text-white font-black text-2xl flex items-center justify-center shadow-md">
                B
              </div>
              <div>
                <div class="text-xs text-slate-500 font-medium">Negosiator Ditugaskan:</div>
                <h3 id="res-assigned" class="text-2xl sm:text-3xl font-black text-slate-900">BONITA</h3>
                <p id="res-reason" class="text-xs sm:text-sm text-slate-600 mt-1 font-medium leading-relaxed">
                  Stronghold Historis Bonita di Kota Bandung. Diprioritaskan untuk menghasilkan nominal diskon terbesar.
                </p>
              </div>
            </div>

            <!-- FAIR WORKLOAD SCORE COMPARISON -->
            <div id="fair-scoring-container" class="mt-4 pt-3.5 border-t border-slate-100">
              <div class="flex flex-col sm:flex-row sm:items-center justify-between text-xs font-bold text-slate-700 mb-2 gap-1">
                <span class="flex items-center text-slate-800">
                  <i class="fa-solid fa-scale-balanced text-blue-600 mr-1.5"></i>Skor Keadilan Beban & Performa (MCDA):
                </span>
                <span class="text-[11px] text-slate-500 font-medium">Bobot: Workload 45% (Prioritas) • Wilayah 25% • Nilai 15% • Speed 15%</span>
              </div>
              <div class="grid grid-cols-2 gap-3 text-xs">
                <div id="score-card-b" class="p-2.5 rounded-xl bg-blue-50/70 border border-blue-200">
                  <div class="flex justify-between font-bold text-blue-900 text-xs">
                    <span>Bonita:</span> <span id="score-val-b" class="font-black">80.0 Poin</span>
                  </div>
                  <div class="w-full bg-blue-200 h-1.5 rounded-full mt-1.5 overflow-hidden">
                    <div id="bar-score-b" class="bg-blue-600 h-full rounded-full transition-all duration-500" style="width: 80%"></div>
                  </div>
                </div>
                <div id="score-card-m" class="p-2.5 rounded-xl bg-teal-50/70 border border-teal-200">
                  <div class="flex justify-between font-bold text-teal-900 text-xs">
                    <span>Mirza:</span> <span id="score-val-m" class="font-black">69.8 Poin</span>
                  </div>
                  <div class="w-full bg-teal-200 h-1.5 rounded-full mt-1.5 overflow-hidden">
                    <div id="bar-score-m" class="bg-teal-600 h-full rounded-full transition-all duration-500" style="width: 70%"></div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <!-- 4 KEY FINANCIAL & SLA METRIC CARDS -->
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <!-- Card 1: Saving -->
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <div class="flex justify-between items-start">
                <span class="text-xs font-bold text-slate-500 uppercase">Potensi Penghematan Sewa</span>
                <span class="p-1.5 rounded-lg bg-emerald-100 text-emerald-700 text-xs"><i class="fa-solid fa-piggy-bank"></i></span>
              </div>
              <div id="res-saving" class="text-2xl font-black text-emerald-600 mt-2">Rp 16.875.000</div>
              <div class="flex items-center space-x-1.5 mt-1">
                <span class="text-xs font-bold text-slate-700">Target Diskon:</span>
                <span id="res-diskon-pct" class="text-xs font-black text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">22.50%</span>
              </div>
            </div>

            <!-- Card 2: Net Price -->
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <div class="flex justify-between items-start">
                <span class="text-xs font-bold text-slate-500 uppercase">Target Harga Net Final</span>
                <span class="p-1.5 rounded-lg bg-blue-100 text-blue-700 text-xs"><i class="fa-solid fa-tag"></i></span>
              </div>
              <div id="res-net-price" class="text-2xl font-black text-slate-900 mt-2">Rp 58.125.000</div>
              <div class="text-xs text-slate-500 mt-1">Batas plafon kesepakatan direksi</div>
            </div>

            <!-- Card 3: SLA Duration -->
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <div class="flex justify-between items-start">
                <span class="text-xs font-bold text-slate-500 uppercase">Estimasi Waktu Closing</span>
                <span class="p-1.5 rounded-lg bg-purple-100 text-purple-700 text-xs"><i class="fa-solid fa-stopwatch"></i></span>
              </div>
              <div class="flex items-baseline space-x-2 mt-2">
                <span id="res-durasi" class="text-2xl font-black text-purple-700">11.2 Hari</span>
                <span class="text-xs text-slate-500">(Target SLA: 17 hr)</span>
              </div>
              <div id="res-sla-status" class="text-xs font-bold text-emerald-700 mt-1">
                <i class="fa-solid fa-circle-check mr-1"></i>SESUAI TARGET SLA
              </div>
            </div>

            <!-- Card 4: Ease Score -->
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
              <div class="flex justify-between items-start">
                <span class="text-xs font-bold text-slate-500 uppercase">Tingkat Kemudahan Negosiasi</span>
                <span class="p-1.5 rounded-lg bg-amber-100 text-amber-700 text-xs"><i class="fa-solid fa-gauge-high"></i></span>
              </div>
              <div id="res-skor" class="text-2xl font-black text-slate-900 mt-2">78.5 <span class="text-xs text-slate-500 font-normal">/ 100</span></div>
              <div id="res-kategori" class="text-xs font-bold text-blue-700 mt-1">MUDAH (Tinggi)</div>
            </div>
          </div>

          <!-- BENCHMARK SIMILAR CITIES CARD -->
          <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
            <div class="text-xs font-bold text-slate-500 uppercase mb-2">Kota / Kabupaten Acuan Pembanding (Benchmark Sejenis)</div>
            <div id="res-similar" class="flex flex-wrap gap-2">
              <span class="px-3 py-1 bg-slate-100 text-slate-800 rounded-lg text-xs font-semibold border border-slate-200">Kota Cimahi</span>
              <span class="px-3 py-1 bg-slate-100 text-slate-800 rounded-lg text-xs font-semibold border border-slate-200">Kota Tasikmalaya</span>
              <span class="px-3 py-1 bg-slate-100 text-slate-800 rounded-lg text-xs font-semibold border border-slate-200">Kota Sukabumi</span>
            </div>
          </div>

        </div>
      </div>
    </div>

    <!-- TAB 2: MULTI RUKO BATCH -->
    <div id="tab-batch" class="tab-content hidden">
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <h2 class="text-lg font-bold text-slate-900 mb-2">Simulasi Evaluasi Beberapa Ruko Sekaligus (Batch Evaluator)</h2>
        <p class="text-xs text-slate-500 mb-4">Gunakan format pemisah pipa (<code class="bg-slate-100 px-1 py-0.5 rounded text-blue-700">|</code>) atau baris baru untuk memasukkan beberapa properti ruko sekaligus.</p>

        <textarea id="batch-input" rows="4" class="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-sm font-mono text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white mb-3"
>Kota Bandung, 75000000
Kab. Bekasi, 45000000
Kota Denpasar, 60000000
Kab. Cianjur, 40000000</textarea>

        <button onclick="runBatchSimulation()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-5 rounded-xl text-sm shadow transition">
          <i class="fa-solid fa-play mr-1.5"></i> Evaluasi Seluruh Ruko
        </button>

        <div id="batch-results-container" class="mt-6 hidden">
          <div class="bg-blue-50 border border-blue-200 rounded-xl p-4 mb-4 flex justify-between items-center">
            <div>
              <div class="text-xs text-blue-700 font-semibold uppercase">Total Proyeksi Penghematan Sewa Batch:</div>
              <div id="batch-total-saving" class="text-2xl font-black text-blue-900 mt-0.5">Rp 34.500.000</div>
            </div>
            <div class="text-right">
              <div class="text-xs text-blue-700 font-semibold uppercase">Total Penawaran Awal:</div>
              <div id="batch-total-asking" class="text-lg font-bold text-slate-800 mt-0.5">Rp 220.000.000</div>
            </div>
          </div>

          <div class="overflow-x-auto">
            <table class="w-full text-left text-sm border border-slate-200 rounded-xl overflow-hidden">
              <thead class="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-600 uppercase">
                <tr>
                  <th class="py-2.5 px-3">No</th>
                  <th class="py-2.5 px-3">Wilayah</th>
                  <th class="py-2.5 px-3">Harga Awal</th>
                  <th class="py-2.5 px-3">Negosiator</th>
                  <th class="py-2.5 px-3">Target Diskon</th>
                  <th class="py-2.5 px-3">Proyeksi Hemat</th>
                  <th class="py-2.5 px-3">Harga Net</th>
                </tr>
              </thead>
              <tbody id="batch-table-body" class="divide-y divide-slate-200 bg-white">
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- TAB 3: MATRIKS KOMPARASI BONITA VS MIRZA -->
    <div id="tab-kpi" class="tab-content hidden space-y-6">
      <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
        
        <!-- Bonita Card -->
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 relative overflow-hidden">
          <div class="flex items-center space-x-3 mb-4">
            <div class="w-12 h-12 rounded-xl bg-blue-600 text-white font-black text-xl flex items-center justify-center shadow">B</div>
            <div>
              <h3 class="text-lg font-bold text-slate-900">BONITA</h3>
              <span class="inline-block px-2.5 py-0.5 bg-blue-50 text-blue-700 text-xs font-bold rounded border border-blue-200">The High-Value Negotiator</span>
            </div>
          </div>
          <div class="grid grid-cols-2 gap-3 mb-4">
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Total Volume Deal</div>
              <div class="text-xl font-bold text-slate-900 mt-0.5">{{ summary_stats.bonita.deal }} <span class="text-xs text-slate-500 font-normal">Cabang</span></div>
            </div>
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Total Akumulasi Saving</div>
              <div class="text-xl font-black text-blue-700 mt-0.5">{{ summary_stats.bonita.saving_fmt }}</div>
            </div>
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Rerata Diskon Sewa</div>
              <div class="text-xl font-black text-emerald-600 mt-0.5">{{ summary_stats.bonita.avg_diskon_fmt }}</div>
            </div>
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Rerata Kecepatan Closing</div>
              <div class="text-xl font-bold text-slate-900 mt-0.5">{{ summary_stats.bonita.avg_durasi_fmt }}</div>
            </div>
          </div>
          <p class="text-xs text-slate-600 leading-relaxed">
            <strong>Karakteristik & Penugasan:</strong> Terbukti unggul signifikan pada ruko harga tinggi (> Rp 50 Juta) dan wilayah prime perkotaan dengan tingkat efisiensi diskon mencapai <strong>{{ summary_stats.bonita.avg_diskon_fmt }}</strong>.
          </p>
        </div>

        <!-- Mirza Card -->
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 relative overflow-hidden">
          <div class="flex items-center space-x-3 mb-4">
            <div class="w-12 h-12 rounded-xl bg-teal-600 text-white font-black text-xl flex items-center justify-center shadow">M</div>
            <div>
              <h3 class="text-lg font-bold text-slate-900">MIRZA</h3>
              <span class="inline-block px-2.5 py-0.5 bg-teal-50 text-teal-700 text-xs font-bold rounded border border-teal-200">The Speed & Volume Specialist</span>
            </div>
          </div>
          <div class="grid grid-cols-2 gap-3 mb-4">
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Total Volume Deal</div>
              <div class="text-xl font-bold text-slate-900 mt-0.5">{{ summary_stats.mirza.deal }} <span class="text-xs text-slate-500 font-normal">Cabang</span></div>
            </div>
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Total Akumulasi Saving</div>
              <div class="text-xl font-black text-teal-700 mt-0.5">{{ summary_stats.mirza.saving_fmt }}</div>
            </div>
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Rerata Diskon Sewa</div>
              <div class="text-xl font-black text-slate-900 mt-0.5">{{ summary_stats.mirza.avg_diskon_fmt }}</div>
            </div>
            <div class="bg-slate-50 p-3 rounded-xl border border-slate-100">
              <div class="text-xs text-slate-500 font-medium">Rerata Kecepatan Closing</div>
              <div class="text-xl font-black text-emerald-600 mt-0.5">{{ summary_stats.mirza.avg_durasi_fmt }}</div>
            </div>
          </div>
          <p class="text-xs text-slate-600 leading-relaxed">
            <strong>Karakteristik & Penugasan:</strong> Terbukti konsisten dan tangguh dalam menangani volume tinggi dengan capaian <strong>{{ summary_stats.mirza.deal }} cabang</strong> dan kecepatan closing teratur di bawah SLA 14 hari kerja.
          </p>
        </div>

      </div>

      <!-- Mann-Whitney U Proof Banner -->
      <div class="bg-slate-900 text-white p-5 rounded-2xl shadow">
        <div class="flex items-center space-x-2 text-xs font-bold text-blue-400 uppercase tracking-wider mb-1">
          <i class="fa-solid fa-square-root-variable"></i>
          <span>Validasi Ilmiah: Uji Statistik Mann-Whitney U</span>
        </div>
        <div class="text-sm font-semibold text-slate-100">
          Statistik U = 11.083,0 | p-value = 0,00012 (Tolak H0 pada Signifikansi 99,9%).
        </div>
        <p class="text-xs text-slate-400 mt-1">
          Perbedaan kinerja antara Bonita dan Mirza bukan fluktuasi acak, melainkan perbedaan persona alami yang wajib dimanfaatkan manajemen melalui sistem penugasan berbasis spesialisasi.
        </p>
      </div>
    </div>

    <!-- TAB: PETA GEOSPASIAL TOPOJSON WILAYAH NEGOSIASI -->
    <div id="tab-geospasial" class="tab-content hidden space-y-6">
      
      <!-- HERO HEADER -->
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div class="flex items-center space-x-2">
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-700 border border-blue-200">
                <i class="fa-solid fa-layer-group mr-1"></i>TopoJSON BPS 524 Kab/Kota
              </span>
              <span class="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-700 border border-emerald-200">
                <i class="fa-solid fa-circle-check mr-1"></i>100% Terpetakan (Zero Unmapped)
              </span>
            </div>
            <h2 class="text-xl font-black text-slate-900 mt-2">
              Peta Geospasial Tingkat Kemudahan Negosiasi Wilayah
            </h2>
            <p class="text-xs text-slate-500 mt-1 max-w-3xl leading-relaxed">
              Pemetaan spasial 76 kabupaten/kota aktif pembukaan cabang PGI (2024–2026) ke dalam batas administratif resmi TopoJSON BPS. Indeks Kemudahan (Skor 0–100) mengintegrasikan dua pilar seimbang: <strong>Kecepatan Closing SLA (50%)</strong> dan <strong>Persentase Efisiensi Diskon Sewa (50%)</strong>.
            </p>
          </div>
          <div class="flex items-center space-x-2 flex-shrink-0">
            <a href="/static/grafik/5_peta_geospasial_kemudahan_nego_2026.png" target="_blank" class="px-4 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition shadow-sm flex items-center space-x-1.5">
              <i class="fa-solid fa-arrow-up-right-from-square"></i>
              <span>Buka Gambar Peta (300 DPI)</span>
            </a>
          </div>
        </div>
      </div>

      <!-- COLOR LEGEND / PANDUAN WARNA KELAS KEMUDAHAN -->
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <!-- Hijau -->
        <div class="bg-white p-4 rounded-xl shadow-sm border-2 border-emerald-500/40 relative overflow-hidden">
          <div class="flex items-center space-x-2.5 mb-2">
            <span class="w-4 h-4 rounded-full bg-emerald-500 flex-shrink-0 shadow-sm"></span>
            <span class="text-xs font-black text-emerald-800 uppercase tracking-wider">Hijau Zamrud</span>
          </div>
          <div class="text-base font-bold text-slate-900">Sangat Mudah (Skor ≥ 70)</div>
          <p class="text-xs text-slate-600 mt-1.5 leading-relaxed">
            <strong>Karakteristik:</strong> Pemilik sangat kooperatif, proses tawar-menawar kilat (rata-rata 3–6 hari kerja), serta rerata diskon sewa tinggi (&gt; 18%).
          </p>
          <div class="mt-2.5 pt-2 border-t border-slate-100 text-[11px] text-emerald-700 font-semibold">
            Contoh: Cilacap (92,2), Tasikmalaya (76,7), Serang (75,2), Jakpus (73,0).
          </div>
        </div>

        <!-- Biru -->
        <div class="bg-white p-4 rounded-xl shadow-sm border-2 border-sky-500/40 relative overflow-hidden">
          <div class="flex items-center space-x-2.5 mb-2">
            <span class="w-4 h-4 rounded-full bg-sky-500 flex-shrink-0 shadow-sm"></span>
            <span class="text-xs font-black text-sky-800 uppercase tracking-wider">Biru Langit</span>
          </div>
          <div class="text-base font-bold text-slate-900">Mudah (Skor 60 – 69,9)</div>
          <p class="text-xs text-slate-600 mt-1.5 leading-relaxed">
            <strong>Karakteristik:</strong> Pasar komersial likuid dan dinamis, durasi tawar-menawar sesuai target SLA (8–12 hari), perolehan diskon stabil (14%–18%).
          </p>
          <div class="mt-2.5 pt-2 border-t border-slate-100 text-[11px] text-sky-700 font-semibold">
            Contoh: Badung Bali (68,4), Kab. Bekasi (67,8), Jaksel (67,6), Surabaya (65,3).
          </div>
        </div>

        <!-- Kuning/Oranye -->
        <div class="bg-white p-4 rounded-xl shadow-sm border-2 border-amber-500/40 relative overflow-hidden">
          <div class="flex items-center space-x-2.5 mb-2">
            <span class="w-4 h-4 rounded-full bg-amber-500 flex-shrink-0 shadow-sm"></span>
            <span class="text-xs font-black text-amber-800 uppercase tracking-wider">Kuning / Oranye</span>
          </div>
          <div class="text-base font-bold text-slate-900">Moderat s/d Sulit (40 – 59,9)</div>
          <p class="text-xs text-slate-600 mt-1.5 leading-relaxed">
            <strong>Karakteristik:</strong> Negosiasi alot dan memakan waktu (12–18 hari kerja), pemilik cenderung menolak diskon besar, diskon moderat (8%–14%).
          </p>
          <div class="mt-2.5 pt-2 border-t border-slate-100 text-[11px] text-amber-700 font-semibold">
            Contoh: Klaten (45,7), Solo (45,3), Purwakarta (43,8), Banyumas (43,0).
          </div>
        </div>

        <!-- Merah -->
        <div class="bg-white p-4 rounded-xl shadow-sm border-2 border-rose-500/40 relative overflow-hidden">
          <div class="flex items-center space-x-2.5 mb-2">
            <span class="w-4 h-4 rounded-full bg-rose-600 flex-shrink-0 shadow-sm"></span>
            <span class="text-xs font-black text-rose-800 uppercase tracking-wider">Merah Tua</span>
          </div>
          <div class="text-base font-bold text-slate-900">Sangat Sulit / Alot (Skor &lt; 40)</div>
          <p class="text-xs text-slate-600 mt-1.5 leading-relaxed">
            <strong>Karakteristik:</strong> Zona Bahaya. Durasi membengkak (&gt; 20 hari), diskon sangat minim (&lt; 10%), risiko gagal diskon melonjak 27,5% (cost of delay).
          </p>
          <div class="mt-2.5 pt-2 border-t border-slate-100 text-[11px] text-rose-700 font-semibold">
            Contoh: Bandung Barat (19,0), Grobogan (20,0), Cirebon Kota (28,3), Cianjur (28,5).
          </div>
        </div>
      </div>

      <!-- GAMBAR PETA TOPOJSON UTAMA -->
      <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4 pb-3 border-b border-slate-100">
          <div class="flex items-center space-x-2">
            <i class="fa-solid fa-map text-blue-600"></i>
            <h3 class="text-sm font-bold text-slate-900 uppercase tracking-wider">Peta Visual Distribusi Spasial Kemudahan Negosiasi (Indonesia, Jawa, Bali & Sumatera)</h3>
          </div>
          <span class="text-xs text-slate-400 font-mono">Batas TopoJSON BPS 524 Wilayah • Resolusi 300 DPI</span>
        </div>
        <div class="overflow-hidden rounded-xl bg-slate-50 border border-slate-200 flex justify-center p-2">
          <img src="/static/grafik/5_peta_geospasial_kemudahan_nego_2026.png" 
               alt="Peta Geospasial TopoJSON Kemudahan Nego PGI" 
               class="w-full object-contain max-h-[650px] rounded-lg shadow-sm">
        </div>
        <div class="mt-3 flex items-center justify-between text-xs text-slate-500">
          <div><i class="fa-solid fa-circle-info text-blue-500 mr-1"></i> Area abu-abu terang pada peta menandakan wilayah administratif BPS yang belum ada ekspansi cabang PGI.</div>
          <a href="/static/grafik/5_peta_geospasial_kemudahan_nego_2026.png" target="_blank" class="font-bold text-blue-600 hover:text-blue-800 flex items-center space-x-1">
            <span>Perbesar Peta Penuh</span> <i class="fa-solid fa-magnifying-glass-plus"></i>
          </a>
        </div>
      </div>

      <!-- GRAFIK PENDUKUNG KARAKTERISTIK & DIKOTOMI -->
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <div class="lg:col-span-7 bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
          <div class="flex items-center justify-between mb-3 pb-2 border-b border-slate-100">
            <h3 class="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center">
              <i class="fa-solid fa-chart-column text-indigo-600 mr-2"></i>Analisis Karakteristik & Top Wilayah
            </h3>
            <a href="/static/grafik/6_analisis_karakteristik_wilayah_nego_2026.png" target="_blank" class="text-xs font-semibold text-blue-600 hover:text-blue-800">
              Perbesar <i class="fa-solid fa-arrow-up-right-from-square ml-1"></i>
            </a>
          </div>
          <div class="overflow-hidden rounded-xl bg-slate-50 border border-slate-200 p-2">
            <img src="/static/grafik/6_analisis_karakteristik_wilayah_nego_2026.png" 
                 alt="Karakteristik Wilayah Nego" 
                 class="w-full object-contain max-h-[420px] rounded-lg">
          </div>
        </div>

        <!-- INSIGHT STRATEGIS KULTURAL UNTUK ATASAN -->
        <div class="lg:col-span-5 bg-white p-5 rounded-2xl shadow-sm border border-slate-200 flex flex-col justify-between">
          <div>
            <h3 class="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center mb-3 pb-2 border-b border-slate-100">
              <i class="fa-solid fa-lightbulb text-amber-500 mr-2"></i>Temuan Kultural untuk Direksi
            </h3>
            <div class="space-y-3 text-xs text-slate-700 leading-relaxed">
              <div class="p-3.5 bg-rose-50/70 rounded-xl border border-rose-200">
                <span class="font-bold text-rose-800 flex items-center">
                  <i class="fa-solid fa-triangle-exclamation mr-1.5"></i>1. Dikotomi Kabupaten vs Kota:
                </span>
                <p class="mt-1 text-slate-700">
                  <strong>8 dari 10 daerah paling alot di Indonesia berstatus KABUPATEN</strong> (KBB, Grobogan, Cianjur, Cirebon, Ciamis, Blora). Ruko di kabupaten didominasi <em>aset warisan keluarga majemuk</em> yang memerlukan konsensus banyak ahli waris di luar kota.
                </p>
              </div>
              <div class="p-3.5 bg-blue-50/70 rounded-xl border border-blue-200">
                <span class="font-bold text-blue-800 flex items-center">
                  <i class="fa-solid fa-city mr-1.5"></i>2. Anomali Saudagar Kota Pusaka:
                </span>
                <p class="mt-1 text-slate-700">
                  Wilayah Kota tidak otomatis mudah. <strong>Kota Cirebon</strong> (skor 28,3) dan <strong>Kota Solo</strong> (skor 45,3) dihuni saudagar lama dengan <em>holding power</em> modal kas kuat yang pantang menurunkan harga sewa.
                </p>
              </div>
              <div class="p-3.5 bg-emerald-50/70 rounded-xl border border-emerald-200">
                <span class="font-bold text-emerald-800 flex items-center">
                  <i class="fa-solid fa-shield-halved mr-1.5"></i>3. Tindakan Manajerial:
                </span>
                <p class="mt-1 text-slate-700">
                  Untuk Zona Merah & Kuning: Tugaskan <strong>Bonita</strong> jika nilai sewa ruko tinggi. Terapkan <strong>cut-off ketat 14 hari kerja</strong> agar tidak membuang biaya sewa berjalan (<em>cost of delay</em>).
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- LIVE SEARCHABLE REGIONAL TABLE -->
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-100">
          <div>
            <h3 class="text-base font-bold text-slate-900">Daftar Skor Kemudahan Seluruh 76 Kabupaten / Kota</h3>
            <p class="text-xs text-slate-500 mt-0.5">Ketik nama wilayah untuk memfilter dan mengecek skor serta kategori warnanya secara instan.</p>
          </div>
          <div class="relative w-full sm:w-72">
            <input type="text" id="table-search-geo" placeholder="Cari nama kota/kabupaten..." 
                   class="w-full bg-slate-50 border border-slate-300 rounded-xl pl-9 pr-3.5 py-2 text-xs font-medium text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white transition"
                   onkeyup="filterGeoTable()">
            <div class="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400 text-xs">
              <i class="fa-solid fa-magnifying-glass"></i>
            </div>
          </div>
        </div>

        <div class="overflow-x-auto max-h-96 overflow-y-auto border border-slate-200 rounded-xl">
          <table class="w-full text-left text-xs" id="geo-table">
            <thead class="bg-slate-100 sticky top-0 z-10 border-b border-slate-200 font-bold text-slate-700 uppercase">
              <tr>
                <th class="py-2.5 px-3">No</th>
                <th class="py-2.5 px-3">Kabupaten / Kota</th>
                <th class="py-2.5 px-3 text-center">Indikator Warna</th>
                <th class="py-2.5 px-3 text-right">Skor Kemudahan</th>
                <th class="py-2.5 px-3 text-center">Kategori</th>
                <th class="py-2.5 px-3 text-right">Rerata Diskon</th>
                <th class="py-2.5 px-3 text-right">Rerata Durasi</th>
                <th class="py-2.5 px-3 text-center">Volume Deal</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100 bg-white">
              {% for row in regional_list %}
              <tr class="geo-row hover:bg-slate-50 transition">
                <td class="py-2 px-3 text-slate-400 font-mono">{{ loop.index }}</td>
                <td class="py-2 px-3 font-bold text-slate-800 geo-name">{{ row.wilayah }}</td>
                <td class="py-2 px-3 text-center">
                  {% if row.skor_kemudahan >= 70 %}
                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                      <span class="w-2 h-2 rounded-full bg-emerald-500 mr-1.5"></span> Hijau Zamrud
                    </span>
                  {% elif row.skor_kemudahan >= 60 %}
                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-sky-100 text-sky-800 border border-sky-300">
                      <span class="w-2 h-2 rounded-full bg-sky-500 mr-1.5"></span> Biru Langit
                    </span>
                  {% elif row.skor_kemudahan >= 40 %}
                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300">
                      <span class="w-2 h-2 rounded-full bg-amber-500 mr-1.5"></span> Kuning / Oranye
                    </span>
                  {% else %}
                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-rose-100 text-rose-800 border border-rose-300">
                      <span class="w-2 h-2 rounded-full bg-rose-600 mr-1.5"></span> Merah Tua
                    </span>
                  {% endif %}
                </td>
                <td class="py-2 px-3 text-right font-black text-slate-900">{{ "%.1f"|format(row.skor_kemudahan) }}</td>
                <td class="py-2 px-3 text-center">
                  <span class="font-semibold text-slate-700">{{ row.kategori }}</span>
                </td>
                <td class="py-2 px-3 text-right font-bold text-emerald-600">{{ "%.1f"|format(row.avg_diskon) }}%</td>
                <td class="py-2 px-3 text-right font-semibold text-slate-700">{{ "%.1f"|format(row.avg_durasi) }} hr</td>
                <td class="py-2 px-3 text-center text-slate-500">{{ row.n_transaksi }} cabang</td>
              </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
      </div>

    </div>

    <!-- TAB 4: REST API DOCS -->
    <div id="tab-api" class="tab-content hidden space-y-6">
      <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
        <h2 class="text-lg font-bold text-slate-900 mb-2">Panduan REST API </h2>
        <p class="text-xs text-slate-600 mb-6 leading-relaxed">
          Gunakan endpoint JSON ini untuk mengintegrasikan model prediksi dan routing langsung ke Aplikasi Mobile Surveyor, ERP, atau Google Sheets korporasi.
        </p>

        <!-- Endpoint 1 -->
        <div class="border border-slate-200 rounded-xl overflow-hidden mb-6">
          <div class="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between">
            <div class="flex items-center space-x-2">
              <span class="px-2.5 py-1 bg-blue-600 text-white font-mono text-xs font-bold rounded">POST</span>
              <span class="font-mono text-sm font-semibold text-slate-800">/api/routing</span>
            </div>
            <span class="text-xs text-slate-500">Auto-Routing Penugasan Ruko Baru</span>
          </div>
          <div class="p-4 bg-slate-900 font-mono text-xs text-emerald-400 overflow-x-auto">
            <div class="text-slate-400">// Contoh Request via cURL:</div>
            curl -X POST http://localhost:5050/api/routing \\<br>
            &nbsp;&nbsp;-H "Content-Type: application/json" \\<br>
            &nbsp;&nbsp;-d '{"wilayah": "Kota Bandung", "harga": 75000000, "load_bonita": 10, "load_mirza": 10}'
          </div>
        </div>

        <!-- Endpoint 2 -->
        <div class="border border-slate-200 rounded-xl overflow-hidden mb-6">
          <div class="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between">
            <div class="flex items-center space-x-2">
              <span class="px-2.5 py-1 bg-emerald-600 text-white font-mono text-xs font-bold rounded">POST</span>
              <span class="font-mono text-sm font-semibold text-slate-800">/api/predict</span>
            </div>
            <span class="text-xs text-slate-500">Simulasi Detail Prediksi Diskon & SLA</span>
          </div>
          <div class="p-4 bg-slate-900 font-mono text-xs text-emerald-400 overflow-x-auto">
            <div class="text-slate-400">// Contoh Request via cURL:</div>
            curl -X POST http://localhost:5050/api/predict \\<br>
            &nbsp;&nbsp;-H "Content-Type: application/json" \\<br>
            &nbsp;&nbsp;-d '{"wilayah": "Kota Bandung", "harga": 75000000, "negosiator": "Bonita"}'
          </div>
        </div>

        <!-- Endpoint 3 -->
        <div class="border border-slate-200 rounded-xl overflow-hidden">
          <div class="bg-slate-50 px-4 py-3 border-b border-slate-200 flex items-center justify-between">
            <div class="flex items-center space-x-2">
              <span class="px-2.5 py-1 bg-slate-700 text-white font-mono text-xs font-bold rounded">GET</span>
              <span class="font-mono text-sm font-semibold text-slate-800">/api/stats/overview</span>
            </div>
            <span class="text-xs text-slate-500">Rekapitulasi Metrik Nasional</span>
          </div>
          <div class="p-4 bg-slate-900 font-mono text-xs text-emerald-400 overflow-x-auto">
            curl -X GET http://localhost:5050/api/stats/overview
          </div>
        </div>

      </div>
    </div>

  </main>

  <!-- FOOTER -->
  <footer class="bg-white border-t border-slate-200 py-4 text-center text-xs text-slate-500 mt-auto">
    Pusat Gadai Indonesia (PGI) — Business Development & Operational Data Analyst &copy; 2026. All Rights Reserved.
  </footer>

  <!-- SCRIPT INTERAKTIF JAVASCRIPT -->
  <script>
    function updateLoadVals() {
      document.getElementById('val-load-b').innerText = document.getElementById('load-bonita').value + " Proyek";
      document.getElementById('val-load-m').innerText = document.getElementById('load-mirza').value + " Proyek";
    }

    document.getElementById('input-harga').addEventListener('input', function(e) {
      const val = parseFloat(e.target.value) || 0;
      document.getElementById('harga-terbilang').innerText = "Rp " + val.toLocaleString('id-ID');
    });

    function switchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(el => {
        el.classList.remove('border-blue-600', 'text-blue-600');
        el.classList.add('border-transparent', 'text-slate-600');
      });

      document.getElementById(tabId).classList.remove('hidden');
      const btn = document.getElementById('btn-' + tabId);
      btn.classList.add('border-blue-600', 'text-blue-600');
      btn.classList.remove('border-transparent', 'text-slate-600');
    }

    async function runSimulation() {
      const wilayah = document.getElementById('input-wilayah').value;
      const harga = parseFloat(document.getElementById('input-harga').value) || 50000000;
      const negosiator = document.getElementById('input-negosiator').value;
      const loadB = parseInt(document.getElementById('load-bonita').value) || 10;
      const loadM = parseInt(document.getElementById('load-mirza').value) || 10;

      if (negosiator === 'auto') {
        // Panggil endpoint routing
        try {
          const res = await fetch('/api/routing', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({wilayah: wilayah, harga: harga, load_bonita: loadB, load_mirza: loadM})
          });
          const data = await res.json();
          renderSimulationResult(data, true);
        } catch (e) {
          console.error(e);
        }
      } else {
        // Panggil endpoint predict
        try {
          const res = await fetch('/api/predict', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({wilayah: wilayah, harga: harga, negosiator: negosiator})
          });
          const data = await res.json();
          renderSimulationResult(data, false);
        } catch (e) {
          console.error(e);
        }
      }
    }

    function renderSimulationResult(data, isRouting) {
      const name = isRouting ? data.assigned : data.negosiator;
      document.getElementById('res-assigned').innerText = name.toUpperCase();
      document.getElementById('res-reason').innerText = isRouting ? data.reason : data.basis_data;
      
      const avatar = document.getElementById('assigned-avatar');
      if (name.includes('BONITA')) {
        avatar.innerText = 'B';
        avatar.className = 'w-16 h-16 rounded-2xl bg-blue-600 text-white font-black text-2xl flex items-center justify-center shadow-md';
      } else if (name.includes('MIRZA')) {
        avatar.innerText = 'M';
        avatar.className = 'w-16 h-16 rounded-2xl bg-teal-600 text-white font-black text-2xl flex items-center justify-center shadow-md';
      } else {
        avatar.innerText = 'T';
        avatar.className = 'w-16 h-16 rounded-2xl bg-slate-700 text-white font-black text-2xl flex items-center justify-center shadow-md';
      }

      document.getElementById('badge-registered').innerText = data.is_registered ? 'Terdaftar di Database' : 'Wilayah Baru (Inferensi ML)';
      document.getElementById('res-saving').innerText = data.potensi_penghematan_exact || ('Rp ' + data.potensi_penghematan_rp.toLocaleString('id-ID'));
      document.getElementById('res-diskon-pct').innerText = (isRouting ? data.est_diskon_pct : data.target_diskon_pct) + '%';
      document.getElementById('res-net-price').innerText = data.target_harga_net_fmt || ('Rp ' + data.target_harga_net.toLocaleString('id-ID'));
      
      const dur = isRouting ? data.est_durasi_hari : data.est_durasi_hari;
      document.getElementById('res-durasi').innerText = dur + ' Hari';
      
      const slaEl = document.getElementById('res-sla-status');
      if (data.deviasi_sla_hari <= 0) {
        slaEl.className = 'text-xs font-bold text-emerald-700 mt-1';
        slaEl.innerHTML = '<i class="fa-solid fa-circle-check mr-1"></i>SESUAI TARGET SLA (≤ 17 hr)';
      } else {
        slaEl.className = 'text-xs font-bold text-amber-700 mt-1';
        slaEl.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>POTENSI TERLAMBAT ' + data.deviasi_sla_hari + ' HARI';
      }

      if (data.scores) {
        document.getElementById('fair-scoring-container').classList.remove('hidden');
        document.getElementById('score-val-b').innerText = data.scores.bonita + ' Poin';
        document.getElementById('score-val-m').innerText = data.scores.mirza + ' Poin';
        document.getElementById('bar-score-b').style.width = Math.min(100, Math.max(5, data.scores.bonita)) + '%';
        document.getElementById('bar-score-m').style.width = Math.min(100, Math.max(5, data.scores.mirza)) + '%';
      } else {
        document.getElementById('fair-scoring-container').classList.add('hidden');
      }

      if (!isRouting && data.skor_komposit) {
        document.getElementById('res-skor').innerHTML = data.skor_komposit + ' <span class="text-xs text-slate-500 font-normal">/ 100</span>';
        document.getElementById('res-kategori').innerText = data.kategori_kemudahan;
      }

      if (data.similar_benchmark) {
        const simContainer = document.getElementById('res-similar');
        simContainer.innerHTML = '';
        data.similar_benchmark.forEach(w => {
          const span = document.createElement('span');
          span.className = 'px-3 py-1 bg-slate-100 text-slate-800 rounded-lg text-xs font-semibold border border-slate-200';
          span.innerText = w;
          simContainer.appendChild(span);
        });
      }
    }

    async function runBatchSimulation() {
      const raw = document.getElementById('batch-input').value.trim();
      const lines = raw.split(/\\r?\\n|\\|/);
      const items = [];

      lines.forEach(l => {
        const parts = l.split(',');
        if (parts.length >= 2) {
          const w = parts[0].trim();
          const h = parseFloat(parts[1].replace(/[^0-9]/g, '')) || 50000000;
          if (w) items.push({wilayah: w, harga: h});
        }
      });

      if (items.length === 0) return;

      try {
        const res = await fetch('/api/batch-evaluate', {
          method: 'POST',
          headers: {'Content-Type': 'application/json'},
          body: JSON.stringify({items: items})
        });
        const data = await res.json();

        document.getElementById('batch-results-container').classList.remove('hidden');
        document.getElementById('batch-total-saving').innerText = data.total_proyeksi_saving_fmt + " (" + data.effective_discount_pct + "%)";
        document.getElementById('batch-total-asking').innerText = data.total_penawaran_fmt;

        const tbody = document.getElementById('batch-table-body');
        tbody.innerHTML = '';

        data.results.forEach((r, idx) => {
          const tr = document.createElement('tr');
          tr.className = 'hover:bg-slate-50';
          const badgeClass = r.assigned === 'BONITA' ? 'bg-blue-100 text-blue-800' : (r.assigned === 'MIRZA' ? 'bg-teal-100 text-teal-800' : 'bg-slate-200 text-slate-800');
          tr.innerHTML = `
            <td class="py-2.5 px-3 font-mono text-xs">${idx + 1}</td>
            <td class="py-2.5 px-3 font-semibold text-slate-900">${r.wilayah_official}</td>
            <td class="py-2.5 px-3 text-slate-600">${r.harga_penawaran_fmt}</td>
            <td class="py-2.5 px-3"><span class="px-2 py-0.5 rounded text-xs font-bold ${badgeClass}">${r.assigned}</span></td>
            <td class="py-2.5 px-3 font-bold text-emerald-600">${r.est_diskon_pct}%</td>
            <td class="py-2.5 px-3 font-bold text-slate-900">${r.potensi_penghematan_exact}</td>
            <td class="py-2.5 px-3 font-medium text-slate-600">${r.target_harga_net_fmt}</td>
          `;
          tbody.appendChild(tr);
        });
      } catch (e) {
        console.error(e);
      }
    }

    function filterGeoTable() {
      const input = document.getElementById('table-search-geo').value.toLowerCase();
      const rows = document.querySelectorAll('.geo-row');
      rows.forEach(r => {
        const name = r.querySelector('.geo-name').innerText.toLowerCase();
        if (name.includes(input)) {
          r.style.display = '';
        } else {
          r.style.display = 'none';
        }
      });
    }

    // Jalankan simulasi default saat pertama kali load
    window.addEventListener('DOMContentLoaded', () => {
      runSimulation();
    });
  </script>
</body>
</html>
"""

@app.route('/')
def index():
    regional_list = REGIONAL_STATS.sort_values(by='skor_kemudahan', ascending=False).to_dict(orient='records')
    return render_template_string(HTML_TEMPLATE, regions=ALL_REGIONS, regional_list=regional_list, summary_stats=SUMMARY_STATS)


# ==============================================================================
# ENTRY POINT & SERVER LAUNCHER
# ==============================================================================
def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"

if __name__ == '__main__':
    port = 5050
    local_ip = get_local_ip()

    print("\n" + "=" * 80)
    print(" 🚀 PUSAT GADAI INDONESIA (PGI) — WEB SERVER & REST API AKTIF")
    print("=" * 80)
    print(f" • Dashboard UI (Browser Laptop) : http://localhost:{port}")
    print(f" • Akses Jaringan Kantor / Wi-Fi : http://{local_ip}:{port}")
    print(f" • Endpoint REST API             : http://localhost:{port}/api/predict")
    print(f" • Status Data                   : {len(DF_GLOBAL)} total cabang | {len(ALL_REGIONS)} wilayah")
    print("-" * 80)
    print(" Tekan Ctrl + C di terminal untuk menghentikan server.")
    print("=" * 80 + "\n")

    app.run(host='0.0.0.0', port=port, debug=False)
