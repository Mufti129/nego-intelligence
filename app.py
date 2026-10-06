#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================================================
PUSAT GADAI INDONESIA (PGI) — EXECUTIVE ANALYTICS & SMART NEGOTIATION INTELLIGENCE
====================================================================================================
Enterprise Decision Support System (DSS) untuk Analisis Performa Negosiasi, Geospasial,
Siklus Lead Time, Manajemen Risiko Vendor, dan Multi-Criteria Decision Analysis (MCDA) Auto-Routing.

Author: Mukhammad Rekza Mufti (Data Analyst — Divisi Bisnis)
====================================================================================================
"""

import os
import sys
import io
import difflib
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy import stats

# ==============================================================================
# 1. KONFIGURASI HALAMAN & ENTERPRISE STYLING
# ==============================================================================
st.set_page_config(
    page_title="PGI Negotiation Intelligence Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Enterprise CSS Styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    /* Main Executive Header */
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
        border: 1px solid #334155;
        border-left: 5px solid #2563eb;
        padding: 22px 26px;
        border-radius: 12px;
        color: #f8fafc;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }
    
    .main-header h1 {
        color: #ffffff !important;
        font-size: 24px !important;
        font-weight: 700 !important;
        margin: 0 !important;
        padding-bottom: 4px !important;
        letter-spacing: -0.3px;
    }
    
    .main-header p {
        color: #94a3b8 !important;
        font-size: 13px !important;
        margin: 0 !important;
        font-weight: 400;
    }
    
    /* KPI Metric Cards */
    .metric-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 16px 18px;
        color: #f8fafc;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.15);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    
    .metric-label {
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #94a3b8;
        margin-bottom: 6px;
    }
    
    .metric-value {
        font-size: 22px;
        font-weight: 700;
        color: #ffffff;
        margin-bottom: 4px;
    }
    
    .metric-sub {
        font-size: 11px;
        color: #38bdf8;
        font-weight: 500;
    }
    
    /* Corporate Badges */
    .badge-primary {
        background: #1e3a8a;
        border: 1px solid #3b82f6;
        border-radius: 10px;
        padding: 16px 20px;
        color: #ffffff;
    }
    
    .badge-success {
        background: #064e3b;
        border: 1px solid #10b981;
        border-radius: 10px;
        padding: 16px 20px;
        color: #ffffff;
    }
    
    .badge-neutral {
        background: #334155;
        border: 1px solid #64748b;
        border-radius: 10px;
        padding: 16px 20px;
        color: #ffffff;
    }

    .tag-status {
        display: inline-block;
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        padding: 3px 8px;
        border-radius: 4px;
        margin-bottom: 8px;
    }
    
    .tag-blue { background: rgba(37, 99, 235, 0.2); color: #60a5fa; border: 1px solid #2563eb; }
    .tag-green { background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid #10b981; }
    .tag-amber { background: rgba(245, 158, 11, 0.2); color: #fbbf24; border: 1px solid #f59e0b; }
    .tag-red { background: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid #ef4444; }
    
    /* Structured Containers */
    .panel-box {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 18px 20px;
        margin-bottom: 16px;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        border-bottom: 1px solid #334155;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0px 0px;
        padding: 8px 16px;
        font-weight: 600;
        font-size: 13px;
        color: #94a3b8;
    }

    .stTabs [aria-selected="true"] {
        color: #38bdf8 !important;
        border-bottom: 2px solid #38bdf8 !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KONFIGURASI GLOBAL & DATASET LOADER
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_NEGO_DIR = os.path.join(BASE_DIR, "data_nego_baru")
FILE_CLEAN_EXCEL = os.path.join(DATA_NEGO_DIR, "data_nego_baru_cleaned_2024-2026.xlsx")
FILE_CLEAN_CSV = os.path.join(BASE_DIR, "properties_cleaned_2024-2026.csv")

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

@st.cache_data(show_spinner=False)
def load_dataset():
    if os.path.exists(FILE_CLEAN_EXCEL):
        df = pd.read_excel(FILE_CLEAN_EXCEL)
    elif os.path.exists(FILE_CLEAN_CSV):
        df = pd.read_csv(FILE_CLEAN_CSV)
    else:
        try:
            import cleansing_data_negosiasi
            df = cleansing_data_negosiasi.run_cleansing(silent=True)
        except:
            df = pd.DataFrame()
            
    if not df.empty:
        if 'nama_negosiator_2' in df.columns:
            df['negosiator'] = df['nama_negosiator_2'].fillna(df.get('nama_negosiator', 'Tanpa Nama')).astype(str).str.strip().str.title()
        elif 'nama_negosiator' in df.columns:
            df['negosiator'] = df['nama_negosiator'].fillna('Tanpa Nama').astype(str).str.strip().str.title()
        else:
            df['negosiator'] = 'Tim Surveyor'
            
        if 'efisiensi_diskon_pct' in df.columns:
            df['diskon_pct'] = pd.to_numeric(df['efisiensi_diskon_pct'], errors='coerce').fillna(0)
        elif 'efisiensi_diskon_nego2' in df.columns:
            df['diskon_pct'] = pd.to_numeric(df['efisiensi_diskon_nego2'], errors='coerce').fillna(0)
        else:
            df['diskon_pct'] = 0.0

        if 'durasi_nego_hari' in df.columns:
            df['durasi_hari'] = pd.to_numeric(df['durasi_nego_hari'], errors='coerce').fillna(0)
        elif 'lama_waktu_realisasi_nego' in df.columns:
            df['durasi_hari'] = pd.to_numeric(df['lama_waktu_realisasi_nego'], errors='coerce').fillna(0)
        else:
            df['durasi_hari'] = 0.0

        if 'diskon_rupiah' in df.columns:
            df['saving_rp'] = pd.to_numeric(df['diskon_rupiah'], errors='coerce').fillna(0)
        else:
            df['saving_rp'] = 0.0

        if 'harga_awal_penawaran' in df.columns:
            df['harga_awal'] = pd.to_numeric(df['harga_awal_penawaran'], errors='coerce').fillna(0)
        elif 'Hargaawal_nego_sheet' in df.columns:
            df['harga_awal'] = pd.to_numeric(df['Hargaawal_nego_sheet'], errors='coerce').fillna(0)
        else:
            df['harga_awal'] = 0.0

        if 'harga_deal_final' in df.columns:
            df['harga_final'] = pd.to_numeric(df['harga_deal_final'], errors='coerce').fillna(0)
        elif 'hargadeal_nego_sheet' in df.columns:
            df['harga_final'] = pd.to_numeric(df['hargadeal_nego_sheet'], errors='coerce').fillna(0)
        else:
            df['harga_final'] = df['harga_awal'] - df['saving_rp']

        if 'wilayah' not in df.columns:
            df['wilayah'] = df.get('kabupaten_kota', 'Wilayah Lainnya')
            
        df['wilayah'] = df['wilayah'].fillna('Wilayah Tidak Terdata').astype(str).str.strip().str.upper()

        if 'Tahun' not in df.columns:
            df['Tahun'] = 2026
            
    return df

DF_GLOBAL = load_dataset()
DF_2026 = DF_GLOBAL[DF_GLOBAL['Tahun'] == 2026].copy() if not DF_GLOBAL.empty and 'Tahun' in DF_GLOBAL.columns else pd.DataFrame()

# ==============================================================================
# 3. STATISTIK REGIONAL & SUMMARY METRICS
# ==============================================================================
@st.cache_data(show_spinner=False)
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

# ==============================================================================
# 4. ENGINE MCDA ROUTING & PREDICTION
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

    b_available = load_bonita < max_cap
    m_available = load_mirza < max_cap

    if not b_available and not m_available:
        assigned = "TIM PENDAMPING (Surveyor: Dika/Salma)"
        reason = (f"Kapasitas kedua negosiator utama penuh (Bonita: {load_bonita}/{max_cap}, "
                  f"Mirza: {load_mirza}/{max_cap} proyek aktif). Rekomendasi eskalasi ke tim pendamping.")
        est_diskon_pct = 12.00
        est_durasi = 14.0
        score_b, score_m = 0.0, 0.0
    elif not b_available:
        assigned = "MIRZA"
        reason = (f"Kapasitas Bonita telah mencapai batas maksimum ({load_bonita}/{max_cap} aktif). "
                  f"Proyek dialihkan ke Mirza untuk menjaga efisiensi target SLA.")
        est_diskon_pct = 15.50 if harga_penawaran >= high_val else 14.00
        est_durasi = 9.8
        score_b, score_m = 0.0, 100.0
    elif not m_available:
        assigned = "BONITA"
        reason = (f"Kapasitas Mirza telah mencapai batas maksimum ({load_mirza}/{max_cap} aktif). "
                  f"Proyek dialihkan ke Bonita untuk menjaga kesinambungan target ekspansi.")
        est_diskon_pct = 19.00 if harga_penawaran >= high_val else 17.50
        est_durasi = 11.5
        score_b, score_m = 100.0, 0.0
    else:
        cap_b = max(0.0, (max_cap - load_bonita) / max_cap * 100.0)
        cap_m = max(0.0, (max_cap - load_mirza) / max_cap * 100.0)

        load_diff = load_bonita - load_mirza
        penalty_b = 0.0
        penalty_m = 0.0
        if load_diff >= 4:
            penalty_b = min(45.0, (load_diff - 3) * 9.0)
        elif load_diff <= -4:
            penalty_m = min(45.0, (abs(load_diff) - 3) * 9.0)

        score_load_b = max(0.0, cap_b - penalty_b)
        score_load_m = max(0.0, cap_m - penalty_m)

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

        if harga_penawaran >= high_val:
            score_val_b = 95.0
            score_val_m = 70.0
        else:
            score_val_b = 75.0
            score_val_m = 90.0

        score_spd_b = 80.0
        score_spd_m = 95.0

        w_load = 0.45
        w_reg = 0.25
        w_val = 0.15
        w_spd = 0.15

        score_b = (w_load * score_load_b) + (w_reg * score_reg_b) + (w_val * score_val_b) + (w_spd * score_spd_b)
        score_m = (w_load * score_load_m) + (w_reg * score_reg_m) + (w_val * score_val_m) + (w_spd * score_spd_m)

        # Guardrail Keadilan Beban Kerja (Maksimal Disparitas 2 Proyek)
        if load_diff >= 2:
            assigned = "MIRZA"
            if harga_penawaran >= high_val:
                reason = (f"Efisiensi Eksekusi: Mirza memiliki rata-rata closing 9,58 hari "
                          f"serta kapasitas aktif prima ({load_mirza} proyek), memastikan realisasi cabang di bawah target SLA.")
            else:
                reason = (f"Akselerasi SLA: Mirza memiliki rekam jejak durasi tercepat "
                          f"(rerata 9,58 hari) serta ketersediaan kapasitas optimal ({load_mirza} proyek aktif).")
            est_diskon_pct = 15.50 if harga_penawaran >= high_val else 14.50
            est_durasi = 9.8
        elif load_diff <= -2:
            assigned = "BONITA"
            if harga_penawaran >= high_val:
                reason = (f"Spesialisasi Nilai Tinggi: Bonita memiliki rekam jejak efisiensi diskon tertinggi "
                          f"(rerata 18,25% dan total saving Rp 1,40 Miliar) dengan kesiapan kapasitas optimal ({load_bonita} proyek aktif).")
            else:
                reason = (f"Efektivitas Diskon: Bonita memiliki rekam jejak yield diskon tinggi "
                          f"(rerata 18,25%) serta kapasitas prima ({load_bonita} proyek aktif).")
            est_diskon_pct = 19.00 if harga_penawaran >= high_val else 17.50
            est_durasi = 11.2
        else:
            if score_b >= score_m:
                assigned = "BONITA"
                if harga_penawaran >= high_val and load_bonita <= load_mirza + 2:
                    reason = (f"Afinitas Nilai Ruko: Bonita unggul dengan skor {score_b:.1f} vs {score_m:.1f}. "
                              f"Kapasitas seimbang ({load_bonita} vs {load_mirza}) dan terbukti paling efektif pada ruko tiket besar (≥ {rupiah(high_val)}).")
                    est_diskon_pct = 19.50
                    est_durasi = 11.6
                elif is_bonita_stronghold:
                    reason = (f"Keunggulan Historis Regional: Bonita unggul dengan skor {score_b:.1f} vs {score_m:.1f}. "
                              f"Stronghold historis di {matched_w} dengan rekam jejak yield diskon tinggi.")
                    est_diskon_pct = 21.00
                    est_durasi = 11.2
                else:
                    reason = (f"Skor Kesesuaian Tertinggi: Bonita ({score_b:.1f}) unggul atas Mirza ({score_m:.1f}) "
                              f"berdasarkan evaluasi ketersediaan kapasitas ({load_bonita} aktif) dan kapabilitas diskon.")
                    est_diskon_pct = 18.00
                    est_durasi = 11.0
            else:
                assigned = "MIRZA"
                if load_mirza < load_bonita:
                    reason = (f"Keseimbangan Beban Tim: Mirza unggul dengan skor {score_m:.1f} vs {score_b:.1f}. "
                              f"Kapasitas Mirza lebih longgar ({load_mirza} aktif vs {load_bonita} aktif) dengan kecepatan closing tinggi.")
                    est_diskon_pct = 15.00 if harga_penawaran >= high_val else 14.50
                    est_durasi = 9.6
                elif is_mirza_stronghold:
                    reason = (f"Keunggulan Historis Regional: Mirza unggul dengan skor {score_m:.1f} vs {score_b:.1f}. "
                              f"Stronghold historis di {matched_w} dengan rekor closing cepat.")
                    est_diskon_pct = 18.00
                    est_durasi = 9.5
                else:
                    reason = (f"Skor Kesesuaian Tertinggi: Mirza ({score_m:.1f}) unggul atas Bonita ({score_b:.1f}) "
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

    neg_clean = negosiator_choice.strip().lower()
    if neg_clean == "bonita":
        mod_diskon = max_diskon * 1.05
        mod_durasi = avg_durasi * 1.02
        assigned_label = "Bonita (High-Value Specialist)"
    elif neg_clean == "mirza":
        mod_diskon = max_diskon * 0.95
        mod_durasi = avg_durasi * 0.88
        assigned_label = "Mirza (Speed Specialist)"
    else:
        mod_diskon = max_diskon
        mod_durasi = avg_durasi
        assigned_label = "Tim Negosiasi Standar"

    potensi_hemat = harga_penawaran * (mod_diskon / 100.0)
    target_net = harga_penawaran - potensi_hemat
    deviasi_sla = mod_durasi - target_sla
    sla_compliance_pct = min(100.0, (target_sla / mod_durasi) * 100.0) if mod_durasi > 0 else 100.0

    subskor_kecepatan = sla_compliance_pct
    subskor_efisiensi = min(100.0, (mod_diskon / 30.0) * 100.0)
    skor_komposit = (0.5 * subskor_kecepatan) + (0.5 * subskor_efisiensi)

    if skor_komposit >= 80: kategori = "Sangat Mudah"
    elif skor_komposit >= 70: kategori = "Mudah"
    elif skor_komposit >= 50: kategori = "Moderat"
    else: kategori = "Sulit (Alot)"

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
# 5. SIDEBAR NAVIGATION & ENTERPRISE CONTROLS
# ==============================================================================
with st.sidebar:
    st.markdown("""
        <div style="padding: 10px 0 18px 0; border-bottom: 1px solid #334155; margin-bottom: 16px;">
            <div style="font-size: 11px; font-weight: 700; color: #38bdf8; letter-spacing: 1px; text-transform: uppercase;">Enterprise DSS</div>
            <div style="font-size: 18px; font-weight: 800; color: #ffffff;">PUSAT GADAI INDONESIA</div>
            <div style="font-size: 12px; color: #94a3b8;">Sistem Analitik Negosiasi UPC</div>
        </div>
    """, unsafe_allow_html=True)

    menu = st.radio(
        "MODUL ANALISIS",
        [
            "Overview & KPI Eksekutif",
            "Smart Auto-Routing & Evaluator AI",
            "Geospatial & Analisis Regional",
            "Siklus Lead Time & SLA",
            "Manajemen Risiko Vendor & Renovasi",
            "Data Explorer & Evaluasi Batch"
        ],
        index=0
    )

    st.markdown("---")
    st.markdown("<div style='font-size: 12px; font-weight: 600; color: #94a3b8; margin-bottom: 8px;'>FILTER TAHUN OPERASIONAL</div>", unsafe_allow_html=True)
    
    available_years = sorted(DF_GLOBAL['Tahun'].dropna().unique().tolist()) if not DF_GLOBAL.empty else [2024, 2025, 2026]
    selected_year = st.selectbox("Periode Data:", ["Semua Tahun"] + [str(y) for y in available_years], index=len(available_years), label_visibility="collapsed")
    
    if selected_year == "Semua Tahun":
        DF_ACTIVE = DF_GLOBAL.copy()
    else:
        DF_ACTIVE = DF_GLOBAL[DF_GLOBAL['Tahun'] == int(selected_year)].copy()

    st.markdown("---")
    st.markdown("""
        <div style="background: #0f172a; padding: 12px 14px; border-radius: 8px; border: 1px solid #334155;">
            <div style="font-size: 11px; color: #94a3b8; line-height: 1.5;">
                <b style="color: #f8fafc;">Status Sistem:</b> Produksi<br>
                <b style="color: #f8fafc;">Basis Data:</b> 1.039 Titik Cabang<br>
                <b style="color: #f8fafc;">Standar SLA:</b> 17.0 Hari Kerja
            </div>
        </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# 6. HEADER UTAMA
# ==============================================================================
st.markdown("""
    <div class="main-header">
        <h1>PUSAT GADAI INDONESIA — NEGOTIATION INTELLIGENCE</h1>
        <p>Sistem Terpadu Optimasi Negosiasi Ruko, Evaluasi Kinerja Regional & Penugasan Berbasis AI Multi-Criteria Decision Analysis (MCDA)</p>
    </div>
""", unsafe_allow_html=True)


# ==============================================================================
# MODUL 1: OVERVIEW & KPI EKSEKUTIF
# ==============================================================================
if menu == "Overview & KPI Eksekutif":
    st.markdown("##### Ringkasan Eksekutif & Key Performance Indicators (KPI)")
    
    tot_db = len(DF_GLOBAL)
    tot_active = len(DF_ACTIVE)
    tot_saving = DF_ACTIVE['saving_rp'].sum()
    avg_diskon = DF_ACTIVE['diskon_pct'].mean() if not DF_ACTIVE.empty else 0
    avg_durasi = DF_ACTIVE['durasi_hari'].mean() if not DF_ACTIVE.empty else 0
    
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Volume Realisasi</div>
                <div class="metric-value">{tot_active:,} <span style="font-size: 13px; font-weight: 500; color: #94a3b8;">Deal</span></div>
                <div class="metric-sub">Database: {tot_db:,} Cabang</div>
            </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Total Efisiensi Biaya</div>
                <div class="metric-value" style="color: #34d399;">{rupiah(tot_saving)}</div>
                <div class="metric-sub">Akumulasi Diskon Riil</div>
            </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Rata-Rata Diskon</div>
                <div class="metric-value" style="color: #38bdf8;">{avg_diskon:.2f}%</div>
                <div class="metric-sub">Target Direksi: ≥ 15.0%</div>
            </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Rata-Rata Durasi Nego</div>
                <div class="metric-value" style="color: #fbbf24;">{avg_durasi:.1f} <span style="font-size: 13px; font-weight: 500; color: #94a3b8;">Hari</span></div>
                <div class="metric-sub">Batas SLA: ≤ 17.0 Hari</div>
            </div>
        """, unsafe_allow_html=True)
    with c5:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Harmonisasi Lead Time</div>
                <div class="metric-value" style="color: #c084fc;">58.7 <span style="font-size: 13px; font-weight: 500; color: #94a3b8;">Hari</span></div>
                <div class="metric-sub">Efisiensi Siklus Cabang</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("##### Evaluasi Komparatif Kinerja Negosiator")
    col_bonita, col_mirza, col_lain = st.columns(3)
    
    df_bonita = DF_ACTIVE[DF_ACTIVE['negosiator'] == 'Bonita']
    df_mirza = DF_ACTIVE[DF_ACTIVE['negosiator'] == 'Mirza']
    df_lain = DF_ACTIVE[~DF_ACTIVE['negosiator'].isin(['Bonita', 'Mirza'])]
    
    with col_bonita:
        st.markdown(f"""
            <div class="metric-card" style="border-left: 3px solid #3b82f6;">
                <div class="tag-status tag-blue">High-Value Specialist</div>
                <div style="font-size: 16px; font-weight: 700; color: #ffffff; margin-bottom: 12px;">BONITA</div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Volume Deal:</span>
                    <b style="color: #f8fafc;">{len(df_bonita)} Deal</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Total Saving:</span>
                    <b style="color: #34d399;">{rupiah(df_bonita['saving_rp'].sum())}</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Rata-rata Diskon:</span>
                    <b style="color: #38bdf8;">{df_bonita['diskon_pct'].mean():.2f}%</b>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 13px;">
                    <span style="color: #94a3b8;">Rata-rata Durasi:</span>
                    <b style="color: #fbbf24;">{df_bonita['durasi_hari'].mean():.1f} Hari</b>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col_mirza:
        st.markdown(f"""
            <div class="metric-card" style="border-left: 3px solid #10b981;">
                <div class="tag-status tag-green">Speed & Volume Specialist</div>
                <div style="font-size: 16px; font-weight: 700; color: #ffffff; margin-bottom: 12px;">MIRZA</div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Volume Deal:</span>
                    <b style="color: #f8fafc;">{len(df_mirza)} Deal</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Total Saving:</span>
                    <b style="color: #34d399;">{rupiah(df_mirza['saving_rp'].sum())}</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Rata-rata Diskon:</span>
                    <b style="color: #38bdf8;">{df_mirza['diskon_pct'].mean():.2f}%</b>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 13px;">
                    <span style="color: #94a3b8;">Rata-rata Durasi:</span>
                    <b style="color: #fbbf24;">{df_mirza['durasi_hari'].mean():.1f} Hari</b>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col_lain:
        st.markdown(f"""
            <div class="metric-card" style="border-left: 3px solid #64748b;">
                <div class="tag-status tag-neutral" style="background: rgba(100, 116, 139, 0.2); color: #cbd5e1; border: 1px solid #64748b;">Surveyor & Lapangan</div>
                <div style="font-size: 16px; font-weight: 700; color: #ffffff; margin-bottom: 12px;">TIM PENDAMPING</div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Volume Deal:</span>
                    <b style="color: #f8fafc;">{len(df_lain)} Deal</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Total Saving:</span>
                    <b style="color: #34d399;">{rupiah(df_lain['saving_rp'].sum())}</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px; font-size: 13px;">
                    <span style="color: #94a3b8;">Rata-rata Diskon:</span>
                    <b style="color: #38bdf8;">{df_lain['diskon_pct'].mean():.2f}%</b>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 13px;">
                    <span style="color: #94a3b8;">Rata-rata Durasi:</span>
                    <b style="color: #fbbf24;">{df_lain['durasi_hari'].mean():.1f} Hari</b>
                </div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c_chart1, c_chart2 = st.columns(2)
    with c_chart1:
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Distribusi Diskon (%) per Negosiator</div>", unsafe_allow_html=True)
        fig_box = px.box(
            DF_ACTIVE[DF_ACTIVE['negosiator'].isin(['Bonita', 'Mirza', 'Tim Surveyor'])],
            x='negosiator',
            y='diskon_pct',
            color='negosiator',
            color_discrete_map={'Bonita': '#3b82f6', 'Mirza': '#10b981', 'Tim Surveyor': '#64748b'},
            points="all",
            labels={'negosiator': 'Negosiator', 'diskon_pct': 'Efisiensi Diskon (%)'},
            template="plotly_dark"
        )
        fig_box.update_layout(showlegend=False, margin=dict(l=20, r=20, t=20, b=20), height=320, plot_bgcolor='#1e293b', paper_bgcolor='#1e293b')
        st.plotly_chart(fig_box, use_container_width=True)

    with c_chart2:
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Harga Awal Penawaran vs Harga Deal Final (Juta Rp)</div>", unsafe_allow_html=True)
        sample_df = DF_ACTIVE[(DF_ACTIVE['harga_awal'] > 0) & (DF_ACTIVE['harga_final'] > 0)].copy()
        sample_df['harga_awal_jt'] = sample_df['harga_awal'] / 1e6
        sample_df['harga_final_jt'] = sample_df['harga_final'] / 1e6
        
        fig_scat = px.scatter(
            sample_df,
            x='harga_awal_jt',
            y='harga_final_jt',
            color='negosiator',
            size='saving_rp',
            hover_data=['wilayah', 'diskon_pct'],
            color_discrete_map={'Bonita': '#3b82f6', 'Mirza': '#10b981'},
            labels={'harga_awal_jt': 'Harga Awal (Juta Rp)', 'harga_final_jt': 'Harga Deal (Juta Rp)'},
            template="plotly_dark"
        )
        fig_scat.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=320, plot_bgcolor='#1e293b', paper_bgcolor='#1e293b')
        st.plotly_chart(fig_scat, use_container_width=True)


# ==============================================================================
# MODUL 2: SMART AUTO-ROUTING & EVALUATOR AI
# ==============================================================================
elif menu == "Smart Auto-Routing & Evaluator AI":
    st.markdown("##### Decision Support System: Evaluasi & Penugasan Negosiator Berbasis AI/MCDA")
    st.markdown("Sistem objektif untuk mengoptimalkan penugasan calon cabang baru berdasarkan analisis beban kerja, keunggulan regional, nilai ruko, dan kecepatan penutupan.")
    
    tab_routing, tab_pred = st.tabs(["Auto-Routing Penugasan (MCDA)", "Estimator Diskon & Indeks Kemudahan"])
    
    with tab_routing:
        col_input, col_result = st.columns([1, 1.2])
        
        with col_input:
            st.markdown("""
                <div class="panel-box">
                    <div style="font-size: 13px; font-weight: 700; color: #38bdf8; text-transform: uppercase; margin-bottom: 12px;">Parameter Calon Gerai</div>
            """, unsafe_allow_html=True)
            
            sim_wilayah = st.selectbox(
                "Wilayah / Kota:",
                options=["KOTA BANDUNG", "KOTA JAKARTA SELATAN", "KAB. BEKASI", "KAB. BREBES", "KOTA SURABAYA", "KAB. TANGERANG", "KOTA SEMARANG", "KAB. KARAWANG", "KOTA DEPOK"] + [w for w in ALL_REGIONS if w not in ["KOTA BANDUNG", "KAB. BEKASI"]],
                index=0,
                key="sb_wil_routing"
            )
            
            sim_harga = st.number_input(
                "Harga Awal Penawaran Pemilik Ruko (Rp/Tahun):",
                min_value=5_000_000,
                max_value=500_000_000,
                value=65_000_000,
                step=5_000_000,
                format="%d",
                key="num_harga_routing"
            )
            st.caption(f"Terbaca: **{rupiah_exact(sim_harga)}** ({rupiah(sim_harga)})")
            
            st.markdown("<div style='margin-top: 14px; font-size: 12px; font-weight: 600; color: #94a3b8;'>Status Beban Kerja Aktif (Maks: 15 Proyek)</div>", unsafe_allow_html=True)
            
            c_b_load, c_m_load = st.columns(2)
            with c_b_load:
                load_b = st.slider("Beban Bonita:", 0, 15, 10, key="sld_b_load")
            with c_m_load:
                load_m = st.slider("Beban Mirza:", 0, 15, 10, key="sld_m_load")
                
            st.markdown("</div>", unsafe_allow_html=True)

        res = execute_routing_logic(sim_wilayah, sim_harga, load_b, load_m)
        
        with col_result:
            assigned_name = res['assigned']
            badge_class = "badge-primary" if "BONITA" in assigned_name else ("badge-success" if "MIRZA" in assigned_name else "badge-neutral")
            
            st.markdown(f"""
                <div class="{badge_class}">
                    <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; opacity: 0.9;">Rekomendasi Penugasan Resmi</div>
                    <div style="font-size: 22px; font-weight: 800; margin: 4px 0 8px 0; color: #ffffff;">{assigned_name}</div>
                    <div style="font-size: 13px; line-height: 1.5; color: #f1f5f9;">
                        <b>Rasional AI:</b> {res['reason']}
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            p1, p2, p3 = st.columns(3)
            with p1:
                st.metric("Potensi Saving", res['potensi_penghematan_fmt'], f"{res['est_diskon_pct']:.2f}% Target Diskon")
            with p2:
                st.metric("Target Harga Net", res['target_harga_net_fmt'], f"-{res['potensi_penghematan_fmt']}")
            with p3:
                st.metric("Estimasi Durasi", f"{res['est_durasi_hari']:.1f} Hari", res['sla_status'])

            st.markdown("<br>", unsafe_allow_html=True)
            
            st.markdown(f"""
                <div style="background: #1e293b; border-radius: 8px; padding: 14px; border: 1px solid #334155;">
                    <div style="font-size: 11px; font-weight: 700; color: #38bdf8; text-transform: uppercase; margin-bottom: 8px;">Skor Komposit MCDA (Bobot: Beban 45% | Regional 25% | Nilai 15% | SLA 15%)</div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; font-size: 13px;">
                        <div>Skor Bonita: <b style="color: #60a5fa; font-size: 15px;">{res['scores']['bonita']}</b> / 100</div>
                        <div>Skor Mirza: <b style="color: #34d399; font-size: 15px;">{res['scores']['mirza']}</b> / 100</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            radar_cats = ['Kapasitas Beban', 'Rekam Jejak Wilayah', 'Afinitas Nilai Ruko', 'Kecepatan SLA']
            cap_b_val = max(0.0, (15 - load_b) / 15 * 100)
            cap_m_val = max(0.0, (15 - load_m) / 15 * 100)
            reg_b_val = 95.0 if any(s in res['wilayah_official'].upper() for s in BONITA_STRONGHOLDS) else 60.0
            reg_m_val = 95.0 if any(s in res['wilayah_official'].upper() for s in MIRZA_STRONGHOLDS) else 60.0
            val_b_val = 95.0 if sim_harga >= 50_000_000 else 75.0
            val_m_val = 70.0 if sim_harga >= 50_000_000 else 90.0
            sla_b_val = 80.0
            sla_m_val = 95.0
            
            fig_rad = go.Figure()
            fig_rad.add_trace(go.Scatterpolar(r=[cap_b_val, reg_b_val, val_b_val, sla_b_val], theta=radar_cats, fill='toself', name='Bonita', line_color='#3b82f6'))
            fig_rad.add_trace(go.Scatterpolar(r=[cap_m_val, reg_m_val, val_m_val, sla_m_val], theta=radar_cats, fill='toself', name='Mirza', line_color='#10b981'))
            fig_rad.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), showlegend=True, template="plotly_dark", margin=dict(l=30, r=30, t=10, b=10), height=230, paper_bgcolor='#1e293b')
            st.plotly_chart(fig_rad, use_container_width=True)

    with tab_pred:
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 12px;'>Estimasi Target Diskon & Klasifikasi Kemudahan Pasar</div>", unsafe_allow_html=True)
        
        c_p_in, c_p_out = st.columns([1, 1.2])
        with c_p_in:
            p_wil = st.selectbox("Wilayah Evaluasi:", options=ALL_REGIONS, index=0, key="sb_wil_pred")
            p_harga = st.number_input("Harga Penawaran Ruko (Rp):", min_value=5_000_000, max_value=500_000_000, value=75_000_000, step=5_000_000, key="num_harga_pred")
            p_neg = st.selectbox("Pilihan Tim Negosiasi:", ["Standar / Tim Gabungan", "Bonita (High-Value Specialist)", "Mirza (Speed Specialist)"], index=0, key="sb_neg_pred")
            
            p_neg_code = "bonita" if "Bonita" in p_neg else ("mirza" if "Mirza" in p_neg else "General")
            pred_res = execute_prediction_simulation(p_wil, p_harga, p_neg_code)

        with c_p_out:
            st.markdown(f"""
                <div class="metric-card" style="border-left: 3px solid #38bdf8;">
                    <div style="font-size: 16px; font-weight: 700; color: #ffffff; margin-bottom: 4px;">Hasil Estimasi: {pred_res['wilayah_official']}</div>
                    <div style="font-size: 12px; color: #94a3b8; margin-bottom: 14px;">Basis Data: {pred_res['basis_data']}</div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 13px;">
                        <div>Target Diskon Maksimal: <b style="color: #34d399;">{pred_res['target_diskon_pct']:.2f}%</b></div>
                        <div>Potensi Penghematan: <b style="color: #34d399;">{pred_res['potensi_penghematan_fmt']}</b></div>
                        <div>Estimasi Durasi: <b>{pred_res['est_durasi_hari']:.1f} Hari</b></div>
                        <div>Target Harga Net: <b>{pred_res['target_harga_net_fmt']}</b></div>
                        <div>Indeks Kemudahan: <b style="color: #fbbf24;">{pred_res['skor_komposit']:.1f} / 100</b></div>
                        <div>Kategori Pasar: <b style="color: #38bdf8;">{pred_res['kategori_kemudahan']}</b></div>
                    </div>
                </div>
            """, unsafe_allow_html=True)


# ==============================================================================
# MODUL 3: GEOSPATIAL & ANALISIS REGIONAL
# ==============================================================================
elif menu == "Geospatial & Analisis Regional":
    st.markdown("##### Pemetaan Spasial Kemudahan Negosiasi Wilayah")
    st.markdown("Klasifikasi efisiensi yield diskon dan kecepatan durasi negosiasi di seluruh Kabupaten/Kota di Indonesia.")
    
    if not REGIONAL_STATS.empty:
        c_kategori_sum = REGIONAL_STATS['kategori'].value_counts()
        
        c_g1, c_g2, c_g3, c_g4 = st.columns(4)
        with c_g1:
            st.metric("Total Wilayah", f"{len(REGIONAL_STATS)} Wilayah")
        with c_g2:
            st.metric("Kategori Sangat Mudah", f"{c_kategori_sum.get('Sangat Mudah', 0)} Wilayah", "Yield Diskon Tinggi")
        with c_g3:
            st.metric("Kategori Moderat", f"{c_kategori_sum.get('Moderat', 0)} Wilayah", "SLA Standar")
        with c_g4:
            st.metric("Kategori Alot / Sulit", f"{c_kategori_sum.get('Sangat Sulit (Alot)', 0) + c_kategori_sum.get('Sulit', 0)} Wilayah", "Perlu Negosiasi Intensif")
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Scatter Matrix: Durasi Negosiasi (Hari) vs Efisiensi Diskon (%)</div>", unsafe_allow_html=True)
        fig_geo_scat = px.scatter(
            REGIONAL_STATS,
            x='avg_durasi',
            y='avg_diskon',
            size='tot_saving',
            color='kategori',
            hover_name='wilayah',
            hover_data={'deal': True, 'avg_asking': ':.0f', 'skor_kemudahan': ':.1f'},
            color_discrete_map={
                'Sangat Mudah': '#10b981',
                'Mudah': '#3b82f6',
                'Moderat': '#fbbf24',
                'Sulit': '#f97316',
                'Sangat Sulit (Alot)': '#ef4444'
            },
            labels={'avg_durasi': 'Rata-rata Durasi (Hari)', 'avg_diskon': 'Rata-rata Diskon (%)'},
            template="plotly_dark"
        )
        fig_geo_scat.add_hline(y=15.0, line_dash="dash", line_color="#94a3b8", annotation_text="Target Diskon 15%")
        fig_geo_scat.add_vline(x=17.0, line_dash="dash", line_color="#f87171", annotation_text="Batas SLA 17 Hari")
        fig_geo_scat.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=400, plot_bgcolor='#1e293b', paper_bgcolor='#1e293b')
        st.plotly_chart(fig_geo_scat, use_container_width=True)

        c_top, c_bot = st.columns(2)
        with c_top:
            st.markdown("<div style='font-size: 13px; font-weight: 600; color: #34d399; margin-bottom: 8px;'>Top 10 Wilayah Paling Mudah (Tinggi Yield & Cepat)</div>", unsafe_allow_html=True)
            top_10 = REGIONAL_STATS.sort_values(by='skor_kemudahan', ascending=False).head(10)
            st.dataframe(
                top_10[['wilayah', 'deal', 'avg_diskon', 'avg_durasi', 'skor_kemudahan', 'kategori']].style.format({
                    'avg_diskon': '{:.2f}%',
                    'avg_durasi': '{:.1f} hr',
                    'skor_kemudahan': '{:.1f}'
                }),
                use_container_width=True,
                hide_index=True
            )
            
        with c_bot:
            st.markdown("<div style='font-size: 13px; font-weight: 600; color: #f87171; margin-bottom: 8px;'>Top 10 Wilayah Paling Sulit (Karakteristik Pasar Alot)</div>", unsafe_allow_html=True)
            bot_10 = REGIONAL_STATS.sort_values(by='skor_kemudahan', ascending=True).head(10)
            st.dataframe(
                bot_10[['wilayah', 'deal', 'avg_diskon', 'avg_durasi', 'skor_kemudahan', 'kategori']].style.format({
                    'avg_diskon': '{:.2f}%',
                    'avg_durasi': '{:.1f} hr',
                    'skor_kemudahan': '{:.1f}'
                }),
                use_container_width=True,
                hide_index=True
            )


# ==============================================================================
# MODUL 4: SIKLUS LEAD TIME & SLA
# ==============================================================================
elif menu == "Siklus Lead Time & SLA":
    st.markdown("##### Dekomposisi 7 Tahapan Siklus Waktu & Harmonisasi Lead Time UPC (2024–2026)")
    st.markdown("Analisis komprehensif 7 tahapan proses pembukaan cabang sejak pengajuan awal formulir survei hingga grand opening resmi.")
    
    # 1. Data Dekomposisi 7 Siklus Tahunan (Identik dengan generate_tabel_siklus.py & pipeline_master)
    df_siklus_tahunan = pd.DataFrame([
        {
            "Periode": "Tahun 2024",
            "Cabang (Unit)": "317 Cabang",
            "Tahap 1: Pengajuan s/d Approved": "N/A*",
            "Tahap 2: Tunggu Nego (Approved ke Nego)": "3.0 hr",
            "Tahap 3: Durasi Nego Riil": "1.4 hr",
            "Tahap 4: Pengumpulan Berkas / TTD": "32.6 hr",
            "Tahap 5: Tunggu Renovasi (TTD ke Mulai)": "3.9 hr",
            "Tahap 6: Durasi Renovasi Fisik": "17.2 hr",
            "Tahap 7: Tunggu Grand Opening": "14.0 hr",
            "Lead Time Pasca-Approval (~58 hr)": "N/A*",
            "Lead Time End-to-End (~69 hr)": "69.5 hr"
        },
        {
            "Periode": "Tahun 2025",
            "Cabang (Unit)": "437 Cabang",
            "Tahap 1: Pengajuan s/d Approved": "13.5 hr",
            "Tahap 2: Tunggu Nego (Approved ke Nego)": "0.0 hr",
            "Tahap 3: Durasi Nego Riil": "9.4 hr",
            "Tahap 4: Pengumpulan Berkas / TTD": "24.4 hr",
            "Tahap 5: Tunggu Renovasi (TTD ke Mulai)": "4.9 hr",
            "Tahap 6: Durasi Renovasi Fisik": "19.9 hr",
            "Tahap 7: Tunggu Grand Opening": "9.3 hr",
            "Lead Time Pasca-Approval (~58 hr)": "57.2 hr",
            "Lead Time End-to-End (~69 hr)": "68.5 hr"
        },
        {
            "Periode": "Tahun 2026",
            "Cabang (Unit)": "285 Cabang",
            "Tahap 1: Pengajuan s/d Approved": "10.7 hr",
            "Tahap 2: Tunggu Nego (Approved ke Nego)": "3.2 hr",
            "Tahap 3: Durasi Nego Riil": "11.5 hr",
            "Tahap 4: Pengumpulan Berkas / TTD": "11.1 hr",
            "Tahap 5: Tunggu Renovasi (TTD ke Mulai)": "7.1 hr",
            "Tahap 6: Durasi Renovasi Fisik": "24.0 hr",
            "Tahap 7: Tunggu Grand Opening": "5.5 hr",
            "Lead Time Pasca-Approval (~58 hr)": "58.4 hr",
            "Lead Time End-to-End (~69 hr)": "69.1 hr"
        },
        {
            "Periode": "Rata-rata Total",
            "Cabang (Unit)": "1.039 Cabang",
            "Tahap 1: Pengajuan s/d Approved": "12.1 hr*",
            "Tahap 2: Tunggu Nego (Approved ke Nego)": "2.2 hr",
            "Tahap 3: Durasi Nego Riil": "7.5 hr",
            "Tahap 4: Pengumpulan Berkas / TTD": "23.2 hr",
            "Tahap 5: Tunggu Renovasi (TTD ke Mulai)": "5.2 hr",
            "Tahap 6: Durasi Renovasi Fisik": "20.2 hr",
            "Tahap 7: Tunggu Grand Opening": "9.7 hr",
            "Lead Time Pasca-Approval (~58 hr)": "58.2 hr*",
            "Lead Time End-to-End (~69 hr)": "69.0 hr"
        }
    ])

    # 4 KPI Summary Cards
    c_s1, c_s2, c_s3, c_s4 = st.columns(4)
    with c_s1:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-label">Lead Time Pasca-Approval</div>
                <div class="metric-value" style="color: #34d399;">58.4 <span style="font-size: 13px; font-weight: 500; color: #94a3b8;">Hari</span></div>
                <div class="metric-sub">Tahap 2 s/d Tahap 7 (~58 Hari)</div>
            </div>
        """, unsafe_allow_html=True)
    with c_s2:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-label">Lead Time End-to-End</div>
                <div class="metric-value" style="color: #c084fc;">69.1 <span style="font-size: 13px; font-weight: 500; color: #94a3b8;">Hari</span></div>
                <div class="metric-sub">Pengajuan Formulir s/d Buka (~69 Hari)</div>
            </div>
        """, unsafe_allow_html=True)
    with c_s3:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-label">Efisiensi Tahap 4 (Berkas)</div>
                <div class="metric-value" style="color: #38bdf8;">-66.0%</div>
                <div class="metric-sub">32.6 hr (2024) -> 11.1 hr (2026)</div>
            </div>
        """, unsafe_allow_html=True)
    with c_s4:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-label">Efisiensi Tahap 7 (Tunggu Buka)</div>
                <div class="metric-value" style="color: #fbbf24;">-60.7%</div>
                <div class="metric-sub">14.0 hr (2024) -> 5.5 hr (2026)</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Visualisasi Stacked Horizontal Bar Chart Dekomposisi 7 Siklus
    st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Dekomposisi 7 Tahapan Siklus Waktu Pembukaan Cabang UPC (Hari Kalender)</div>", unsafe_allow_html=True)
    
    bar_years = ['2024 (317 cabang)', '2025 (437 cabang)', '2026 (285 cabang)']
    fig_siklus = go.Figure()
    
    stages_bars = [
        ('Tahap 1: Pengajuan s/d Approved', [0.0, 13.5, 10.7], '#8B5CF6'),
        ('Tahap 2: Tunggu Nego (Approved ke Nego)', [3.0, 0.0, 3.2], '#94A3B8'),
        ('Tahap 3: Durasi Nego Riil', [1.4, 9.4, 11.5], '#2563EB'),
        ('Tahap 4: Pengumpulan Berkas / TTD', [32.6, 24.4, 11.1], '#06B6D4'),
        ('Tahap 5: Tunggu Renovasi (TTD ke Mulai)', [3.9, 4.9, 7.1], '#64748B'),
        ('Tahap 6: Durasi Renovasi Fisik', [17.2, 19.9, 24.0], '#F59E0B'),
        ('Tahap 7: Tunggu Grand Opening', [14.0, 9.3, 5.5], '#10B981')
    ]
    
    for name, vals, color in stages_bars:
        fig_siklus.add_trace(go.Bar(
            y=bar_years,
            x=vals,
            name=name,
            orientation='h',
            marker_color=color
        ))
        
    fig_siklus.update_layout(
        barmode='stack',
        template="plotly_dark",
        margin=dict(l=20, r=20, t=10, b=20),
        height=320,
        plot_bgcolor='#1e293b',
        paper_bgcolor='#1e293b',
        legend=dict(orientation="h", yanchor="bottom", y=-0.45, xanchor="center", x=0.5, font=dict(size=10)),
        xaxis=dict(title="Rata-rata Durasi (Hari Kalender)")
    )
    st.plotly_chart(fig_siklus, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Multi-Tab: Tabel Tahunan vs Bulanan 2026
    tab_tab_yr, tab_tab_mo = st.tabs(["Tabel Dekomposisi 7 Siklus Tahunan (2024–2026)", "Detail Bulanan 7 Siklus Tahun 2026"])
    
    with tab_tab_yr:
        st.markdown("<div style='font-size: 13px; font-weight: 600; color: #f8fafc; margin-bottom: 8px;'>Tabel Dekomposisi 7 Siklus Waktu Pembukaan Cabang UPC (2024–2026)</div>", unsafe_allow_html=True)
        st.dataframe(df_siklus_tahunan, use_container_width=True, hide_index=True)
        st.caption("Catatan: Tahap 1 pada tahun 2024 bernilai N/A* karena digitalisasi pencatatan tanggal survei baru dibakukan pada 2025.")

    with tab_tab_mo:
        st.markdown("<div style='font-size: 13px; font-weight: 600; color: #f8fafc; margin-bottom: 8px;'>Tabel Detail Bulanan 7 Tahapan Siklus Waktu Tahun 2026 (Januari – Agustus)</div>", unsafe_allow_html=True)
        
        df_siklus_bulanan_2026 = pd.DataFrame([
            {"Periode": "Januari 2026", "Cabang": "27", "Tahap 1: Pengajuan s/d Approved": "12.8 hr", "Tahap 2: Tunggu Nego": "4.2 hr", "Tahap 3: Durasi Nego": "9.2 hr", "Tahap 4: Pengumpulan Berkas": "19.5 hr", "Tahap 5: Tunggu Renovasi": "7.2 hr", "Tahap 6: Durasi Renovasi": "23.4 hr", "Tahap 7: Tunggu Open": "6.8 hr", "Pasca-Approval": "63.7 hr", "End-to-End": "76.5 hr"},
            {"Periode": "Februari 2026", "Cabang": "31", "Tahap 1: Pengajuan s/d Approved": "11.5 hr", "Tahap 2: Tunggu Nego": "3.8 hr", "Tahap 3: Durasi Nego": "10.4 hr", "Tahap 4: Pengumpulan Berkas": "16.2 hr", "Tahap 5: Tunggu Renovasi": "6.8 hr", "Tahap 6: Durasi Renovasi": "24.1 hr", "Tahap 7: Tunggu Open": "6.1 hr", "Pasca-Approval": "60.6 hr", "End-to-End": "72.1 hr"},
            {"Periode": "Maret 2026", "Cabang": "36", "Tahap 1: Pengajuan s/d Approved": "10.9 hr", "Tahap 2: Tunggu Nego": "3.5 hr", "Tahap 3: Durasi Nego": "11.1 hr", "Tahap 4: Pengumpulan Berkas": "13.4 hr", "Tahap 5: Tunggu Renovasi": "7.4 hr", "Tahap 6: Durasi Renovasi": "24.8 hr", "Tahap 7: Tunggu Open": "5.7 hr", "Pasca-Approval": "59.2 hr", "End-to-End": "70.1 hr"},
            {"Periode": "April 2026", "Cabang": "34", "Tahap 1: Pengajuan s/d Approved": "10.2 hr", "Tahap 2: Tunggu Nego": "3.1 hr", "Tahap 3: Durasi Nego": "11.8 hr", "Tahap 4: Pengumpulan Berkas": "11.8 hr", "Tahap 5: Tunggu Renovasi": "7.0 hr", "Tahap 6: Durasi Renovasi": "24.5 hr", "Tahap 7: Tunggu Open": "5.4 hr", "Pasca-Approval": "58.1 hr", "End-to-End": "68.3 hr"},
            {"Periode": "Mei 2026", "Cabang": "39", "Tahap 1: Pengajuan s/d Approved": "9.8 hr", "Tahap 2: Tunggu Nego": "2.9 hr", "Tahap 3: Durasi Nego": "12.2 hr", "Tahap 4: Pengumpulan Berkas": "9.7 hr", "Tahap 5: Tunggu Renovasi": "7.3 hr", "Tahap 6: Durasi Renovasi": "23.9 hr", "Tahap 7: Tunggu Open": "5.2 hr", "Pasca-Approval": "56.2 hr", "End-to-End": "66.0 hr"},
            {"Periode": "Juni 2026", "Cabang": "42", "Tahap 1: Pengajuan s/d Approved": "9.5 hr", "Tahap 2: Tunggu Nego": "2.8 hr", "Tahap 3: Durasi Nego": "12.5 hr", "Tahap 4: Pengumpulan Berkas": "8.4 hr", "Tahap 5: Tunggu Renovasi": "6.9 hr", "Tahap 6: Durasi Renovasi": "23.6 hr", "Tahap 7: Tunggu Open": "4.9 hr", "Pasca-Approval": "54.1 hr", "End-to-End": "63.6 hr"},
            {"Periode": "Juli 2026", "Cabang": "38", "Tahap 1: Pengajuan s/d Approved": "8.9 hr", "Tahap 2: Tunggu Nego": "2.4 hr", "Tahap 3: Durasi Nego": "12.1 hr", "Tahap 4: Pengumpulan Berkas": "5.2 hr", "Tahap 5: Tunggu Renovasi": "6.7 hr", "Tahap 6: Durasi Renovasi": "23.8 hr", "Tahap 7: Tunggu Open": "4.6 hr", "Pasca-Approval": "48.8 hr", "End-to-End": "57.7 hr"},
            {"Periode": "Agustus 2026", "Cabang": "38", "Tahap 1: Pengajuan s/d Approved": "8.1 hr", "Tahap 2: Tunggu Nego": "1.9 hr", "Tahap 3: Durasi Nego": "11.6 hr", "Tahap 4: Pengumpulan Berkas": "3.7 hr", "Tahap 5: Tunggu Renovasi": "6.5 hr", "Tahap 6: Durasi Renovasi": "23.5 hr", "Tahap 7: Tunggu Open": "4.3 hr", "Pasca-Approval": "40.0 hr", "End-to-End": "48.1 hr"},
            {"Periode": "Rata-rata 2026", "Cabang": "285", "Tahap 1: Pengajuan s/d Approved": "10.7 hr", "Tahap 2: Tunggu Nego": "3.2 hr", "Tahap 3: Durasi Nego": "11.5 hr", "Tahap 4: Pengumpulan Berkas": "11.1 hr", "Tahap 5: Tunggu Renovasi": "7.1 hr", "Tahap 6: Durasi Renovasi": "24.0 hr", "Tahap 7: Tunggu Open": "5.5 hr", "Pasca-Approval": "58.4 hr", "End-to-End": "69.1 hr"}
        ])
        st.dataframe(df_siklus_bulanan_2026, use_container_width=True, hide_index=True)

    st.markdown("---")
    
    # Law of Diminishing Returns Section
    st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Evaluasi Law of Diminishing Returns: Durasi Negosiasi vs Efisiensi Diskon</div>", unsafe_allow_html=True)
    
    if not DF_ACTIVE.empty:
        df_valid_dur = DF_ACTIVE[(DF_ACTIVE['durasi_hari'] > 0) & (DF_ACTIVE['durasi_hari'] <= 40)].copy()
        df_valid_dur['durasi_bin'] = pd.cut(
            df_valid_dur['durasi_hari'],
            bins=[0, 5, 10, 15, 20, 100],
            labels=['1-5 Hari (Kilat)', '6-10 Hari (Optimal)', '11-15 Hari (Standar)', '16-20 Hari (Alot)', '> 20 Hari (Inefisiensi)']
        )
        
        dim_summary = df_valid_dur.groupby('durasi_bin', observed=False).agg(
            total_deal=('diskon_pct', 'count'),
            avg_diskon=('diskon_pct', 'mean'),
            avg_saving=('saving_rp', 'mean')
        ).reset_index()
        
        c_dim1, c_dim2 = st.columns(2)
        with c_dim1:
            fig_dim = px.bar(
                dim_summary,
                x='durasi_bin',
                y='avg_diskon',
                color='avg_diskon',
                color_continuous_scale='Blues',
                labels={'durasi_bin': 'Kelompok Durasi', 'avg_diskon': 'Rata-rata Diskon (%)'},
                template="plotly_dark",
                text='avg_diskon'
            )
            fig_dim.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
            fig_dim.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=300, plot_bgcolor='#1e293b', paper_bgcolor='#1e293b')
            st.plotly_chart(fig_dim, use_container_width=True)
            
        with c_dim2:
            st.markdown("""
                <div class="metric-card" style="height: 100%;">
                    <div style="font-size: 14px; font-weight: 700; color: #38bdf8; margin-bottom: 8px;">Rekomendasi Kebijakan Negosiasi & SLA</div>
                    <div style="font-size: 13px; line-height: 1.6; color: #cbd5e1;">
                        1. <b>Rentang Optimal:</b> Rata-rata efisiensi diskon tertinggi terjadi pada rentang <b>1–10 hari</b> (17.8%).<br>
                        2. <b>Diminishing Returns:</b> Perpanjangan waktu di atas 14 hari tidak menghasilkan peningkatan diskon secara signifikan (p-value > 0.05).<br>
                        3. <b>Hard Stop Policy:</b> Direkomendasikan penetapan batas maksimal negosiasi pada hari ke-15 untuk mencegah keterlambatan pembukaan gerai baru.
                    </div>
                </div>
            """, unsafe_allow_html=True)


# ==============================================================================
# MODUL 5: MANAJEMEN RISIKO VENDOR & RENOVASI
# ==============================================================================
elif menu == "Manajemen Risiko Vendor & Renovasi":
    st.markdown("##### Evaluasi Risiko Kontraktor Renovasi")
    st.markdown("Analisis kepatuhan deviasi biaya (cost variance) dan ketepatan waktu pengerjaan (lead time).")
    
    vendor_data = [
        {"Kontraktor": "CV Cipta Karya Mandiri", "Total Proyek": 28, "Deviasi Biaya (%)": 2.1, "Keterlambatan (Hari)": 1.5, "Kategori Kinerja": "Tier 1 (Mitra Unggulan)"},
        {"Kontraktor": "PT Bangun Graha Utama", "Total Proyek": 24, "Deviasi Biaya (%)": 3.8, "Keterlambatan (Hari)": 2.2, "Kategori Kinerja": "Tier 1 (Mitra Unggulan)"},
        {"Kontraktor": "CV Sentosa Jaya Konstruksi", "Total Proyek": 19, "Deviasi Biaya (%)": 5.4, "Keterlambatan (Hari)": 4.1, "Kategori Kinerja": "Tier 2 (Moderat)"},
        {"Kontraktor": "CV Prima Jaya Teknik", "Total Proyek": 15, "Deviasi Biaya (%)": 8.2, "Keterlambatan (Hari)": 6.8, "Kategori Kinerja": "Tier 3 (Risiko Tinggi)"},
        {"Kontraktor": "Kontraktor Trimo (Wilayah Bali)", "Total Proyek": 8, "Deviasi Biaya (%)": 18.5, "Keterlambatan (Hari)": 14.5, "Kategori Kinerja": "Red Flag (Kasus Kritis)"},
    ]
    df_vendor = pd.DataFrame(vendor_data)
    
    c_v1, c_v2 = st.columns([1.2, 1])
    with c_v1:
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Kuadran Risiko Keterlambatan vs Deviasi Biaya</div>", unsafe_allow_html=True)
        fig_vend = px.scatter(
            df_vendor,
            x='Keterlambatan (Hari)',
            y='Deviasi Biaya (%)',
            size='Total Proyek',
            color='Kategori Kinerja',
            hover_name='Kontraktor',
            color_discrete_map={
                'Tier 1 (Mitra Unggulan)': '#10b981',
                'Tier 2 (Moderat)': '#3b82f6',
                'Tier 3 (Risiko Tinggi)': '#f97316',
                'Red Flag (Kasus Kritis)': '#ef4444'
            },
            template="plotly_dark"
        )
        fig_vend.add_vline(x=5.0, line_dash="dash", line_color="#fbbf24", annotation_text="Batas Waktu 5 Hari")
        fig_vend.add_hline(y=7.0, line_dash="dash", line_color="#fbbf24", annotation_text="Batas Biaya 7%")
        fig_vend.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=340, plot_bgcolor='#1e293b', paper_bgcolor='#1e293b')
        st.plotly_chart(fig_vend, use_container_width=True)

    with c_v2:
        st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin-bottom: 10px;'>Matriks Kinerja Kontraktor</div>", unsafe_allow_html=True)
        st.dataframe(df_vendor, use_container_width=True, hide_index=True)
        st.warning("Catatan Evaluasi: Kontraktor Trimo (Bali) memerlukan audit kontrak akibat rata-rata keterlambatan 14.5 hari dan deviasi anggaran 18.5%.")


# ==============================================================================
# MODUL 6: DATA EXPLORER & EVALUASI BATCH
# ==============================================================================
elif menu == "Data Explorer & Evaluasi Batch":
    st.markdown("##### Eksplorasi Data Cabang & Evaluasi Batch")
    st.markdown("Pencarian dan penelusuran database 1.039 cabang serta pemrosesan serentak file ruko baru.")
    
    tab_exp, tab_batch = st.tabs(["Eksplorasi Data Cabang", "Evaluasi Batch Multi-Ruko (AI)"])
    
    with tab_exp:
        f_c1, f_c2, f_c3 = st.columns(3)
        with f_c1:
            filter_nego = st.multiselect("Filter Negosiator:", options=sorted(DF_ACTIVE['negosiator'].unique().tolist()), default=None)
        with f_c2:
            filter_wil = st.multiselect("Filter Wilayah:", options=sorted(DF_ACTIVE['wilayah'].unique().tolist()), default=None)
        with f_c3:
            search_query = st.text_input("Pencarian Nama Cabang / No Pengajuan:", "")

        df_filtered = DF_ACTIVE.copy()
        if filter_nego:
            df_filtered = df_filtered[df_filtered['negosiator'].isin(filter_nego)]
        if filter_wil:
            df_filtered = df_filtered[df_filtered['wilayah'].isin(filter_wil)]
        if search_query:
            query = search_query.strip().lower()
            df_filtered = df_filtered[
                df_filtered['nama_cabang'].astype(str).str.lower().str.contains(query, na=False) |
                df_filtered['nomor_pengajuan'].astype(str).str.lower().str.contains(query, na=False) |
                df_filtered['wilayah'].astype(str).str.lower().str.contains(query, na=False)
            ]

        st.markdown(f"<div style='font-size: 12px; color: #94a3b8; margin-bottom: 8px;'>Menampilkan <b>{len(df_filtered):,}</b> dari total <b>{len(DF_ACTIVE):,}</b> baris data</div>", unsafe_allow_html=True)
        
        display_cols = ['nomor_pengajuan', 'nama_cabang', 'wilayah', 'negosiator', 'harga_awal', 'harga_final', 'saving_rp', 'diskon_pct', 'durasi_hari']
        available_display_cols = [c for c in display_cols if c in df_filtered.columns]
        
        st.dataframe(
            df_filtered[available_display_cols].head(500).style.format({
                'harga_awal': lambda x: rupiah(x),
                'harga_final': lambda x: rupiah(x),
                'saving_rp': lambda x: rupiah(x),
                'diskon_pct': '{:.2f}%',
                'durasi_hari': '{:.1f}'
            }),
            use_container_width=True,
            height=380
        )
        
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            csv_buf = df_filtered[available_display_cols].to_csv(index=False).encode('utf-8')
            st.download_button("Ekspor Data Terfilter (CSV)", data=csv_buf, file_name="data_negosiasi_pgi_filtered.csv", mime="text/csv", use_container_width=True)
        with c_dl2:
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
                df_filtered[available_display_cols].to_excel(writer, index=False, sheet_name='Data_Nego')
            st.download_button("Ekspor Data Terfilter (Excel)", data=excel_buf.getvalue(), file_name="data_negosiasi_pgi_filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with tab_batch:
        st.markdown("<div style='font-size: 13px; color: #cbd5e1; margin-bottom: 12px;'>Unggah file spreadsheet (CSV atau Excel) dengan kolom <code>wilayah</code> dan <code>harga_penawaran</code> untuk memproses rekomendasi secara serentak.</div>", unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader("Pilih Berkas CSV / Excel:", type=['csv', 'xlsx', 'xls'], label_visibility="collapsed")
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    up_df = pd.read_csv(uploaded_file)
                else:
                    up_df = pd.read_excel(uploaded_file)
                    
                st.success(f"Berhasil memuat {len(up_df)} data cabang.")
                
                w_col = next((c for c in up_df.columns if 'wilayah' in c.lower() or 'kota' in c.lower()), None)
                p_col = next((c for c in up_df.columns if 'harga' in c.lower() or 'asking' in c.lower() or 'penawaran' in c.lower()), None)
                
                if w_col and p_col:
                    results_batch = []
                    for _, row in up_df.iterrows():
                        w_val = str(row[w_col])
                        p_val = float(pd.to_numeric(row[p_col], errors='coerce') or 50_000_000)
                        eval_res = execute_routing_logic(w_val, p_val)
                        results_batch.append({
                            'Wilayah Input': w_val,
                            'Wilayah Resmi': eval_res['wilayah_official'],
                            'Harga Penawaran': p_val,
                            'Rekomendasi Negosiator': eval_res['assigned'],
                            'Rasional AI': eval_res['reason'],
                            'Estimasi Diskon (%)': eval_res['est_diskon_pct'],
                            'Potensi Saving (Rp)': eval_res['potensi_penghematan_rp'],
                            'Estimasi Durasi (Hari)': eval_res['est_durasi_hari'],
                            'Status SLA': eval_res['sla_status']
                        })
                    
                    batch_res_df = pd.DataFrame(results_batch)
                    st.markdown("<div style='font-size: 14px; font-weight: 600; color: #f8fafc; margin: 16px 0 8px 0;'>Hasil Evaluasi AI Multi-Ruko</div>", unsafe_allow_html=True)
                    st.dataframe(
                        batch_res_df.style.format({
                            'Harga Penawaran': lambda x: rupiah(x),
                            'Potensi Saving (Rp)': lambda x: rupiah(x),
                            'Estimasi Diskon (%)': '{:.2f}%',
                            'Estimasi Durasi (Hari)': '{:.1f}'
                        }),
                        use_container_width=True
                    )
                    
                    b_excel_buf = io.BytesIO()
                    with pd.ExcelWriter(b_excel_buf, engine='xlsxwriter') as writer:
                        batch_res_df.to_excel(writer, index=False, sheet_name='Hasil_Batch_AI')
                    st.download_button("Unduh Hasil Evaluasi AI (Excel)", data=b_excel_buf.getvalue(), file_name="hasil_batch_ai_routing_pgi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("Format kolom tidak sesuai. Pastikan berkas memiliki kolom 'wilayah' dan 'harga_penawaran'.")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses data: {e}")


# ==============================================================================
# FOOTER
# ==============================================================================
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #64748b; font-size: 11px; padding: 6px 0 16px 0;">
        © 2026 <b>Pusat Gadai Indonesia (PGI)</b> — Divisi Bisnis & Ekspansi Jaringan UPC.<br>
        Sistem Analisis & Engine Rekomendasi Terpadu.
    </div>
""", unsafe_allow_html=True)
