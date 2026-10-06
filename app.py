#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
====================================================================================================
PUSAT GADAI INDONESIA (PGI) — EXECUTIVE ANALYTICS & SMART NEGOTIATION DASHBOARD
====================================================================================================
Aplikasi Web Terpadu Streamlit Cloud untuk Analisis Performa Negosiasi, Geospasial,
Siklus Lead Time, Vendor Risk, dan Simulasi Smart Auto-Routing Negosiator Berbasis AI/MCDA.

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
# 1. KONFIGURASI HALAMAN & GLOBAL STYLING
# ==============================================================================
st.set_page_config(
    page_title="PGI Smart Negotiation Intelligence Dashboard",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Modern Styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #0284c7 50%, #0d9488 100%);
        padding: 24px 28px;
        border-radius: 16px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2), 0 8px 10px -6px rgba(0, 0, 0, 0.2);
    }
    
    .main-header h1 {
        color: #ffffff !important;
        font-size: 26px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        padding-bottom: 6px !important;
        letter-spacing: -0.5px;
    }
    
    .main-header p {
        color: #e0f2fe !important;
        font-size: 14px !important;
        margin: 0 !important;
        font-weight: 400;
    }
    
    .metric-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 14px;
        padding: 18px 20px;
        color: #f8fafc;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        border-color: #38bdf8;
    }
    
    .metric-label {
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #94a3b8;
        margin-bottom: 6px;
    }
    
    .metric-value {
        font-size: 24px;
        font-weight: 800;
        color: #f8fafc;
        margin-bottom: 4px;
    }
    
    .metric-sub {
        font-size: 11px;
        color: #38bdf8;
        font-weight: 500;
    }
    
    .hero-badge-bonita {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        border: 2px solid #60a5fa;
        border-radius: 16px;
        padding: 20px;
        color: white;
        box-shadow: 0 10px 15px -3px rgba(59, 130, 246, 0.3);
    }
    
    .hero-badge-mirza {
        background: linear-gradient(135deg, #065f46 0%, #10b981 100%);
        border: 2px solid #34d399;
        border-radius: 16px;
        padding: 20px;
        color: white;
        box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.3);
    }
    
    .hero-badge-tim {
        background: linear-gradient(135deg, #7c2d12 0%, #f97316 100%);
        border: 2px solid #fb923c;
        border-radius: 16px;
        padding: 20px;
        color: white;
        box-shadow: 0 10px 15px -3px rgba(249, 115, 22, 0.3);
    }
    
    .glass-box {
        background: rgba(30, 41, 59, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(8px);
        border-radius: 14px;
        padding: 20px;
        margin-bottom: 16px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0px 0px;
        padding: 10px 18px;
        font-weight: 600;
        font-size: 14px;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KONFIGURASI GLOBAL & DATASET LOADER (IDENTIK DENGAN APP_SERVER_API)
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
    """
    Memuat dataset bersih siap pakai:
    Prioritas utama membaca data_nego_baru/data_nego_baru_cleaned_2024-2026.xlsx.
    Fallback ke properties_cleaned_2024-2026.csv atau kompilasi data_nego_baru.
    """
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
# 3. PRE-KALKULASI STATISTIK WILAYAH & SUMMARY
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
# 4. LOGIKA INFERENSI, MATCHING & SMART ROUTING (IDENTIK PERSIS 100%)
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

    # --- MODEL AI: MULTI-CRITERIA DECISION ANALYSIS (MCDA) BERBASIS KEADILAN BEBAN ---
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
        if harga_penawaran >= high_val:
            score_val_b = 95.0
            score_val_m = 70.0
        else:
            score_val_b = 75.0
            score_val_m = 90.0

        # 4. Pilar 4: Speed & SLA Responsiveness (0 - 100) -> Bobot 15%
        score_spd_b = 80.0
        score_spd_m = 95.0

        # Kalkulasi Skor Komposit
        w_load = 0.45
        w_reg = 0.25
        w_val = 0.15
        w_spd = 0.15

        score_b = (w_load * score_load_b) + (w_reg * score_reg_b) + (w_val * score_val_b) + (w_spd * score_spd_b)
        score_m = (w_load * score_load_m) + (w_reg * score_reg_m) + (w_val * score_val_m) + (w_spd * score_spd_m)

        # GUARDRAIL KEADILAN MUTLAK: BATAS SELISIH PROYEK MAKSIMAL 2
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
# 5. SIDEBAR NAVIGATION & FILTER GLOBAL
# ==============================================================================
with st.sidebar:
    st.markdown("""
        <div style="text-align: center; padding: 10px 0 20px 0;">
            <h2 style="color: #38bdf8; margin: 0; font-weight: 800; font-size: 22px;">🏢 PGI ANALYTICS</h2>
            <p style="color: #94a3b8; font-size: 12px; margin: 2px 0 0 0;">Sistem Cerdas Negosiasi UPC</p>
        </div>
    """, unsafe_allow_html=True)

    menu = st.radio(
        "NAVIGASI MODUL",
        [
            "📊 Executive Dashboard",
            "🤖 Smart AI Auto-Routing & Simulator",
            "🗺️ Geospatial & Regional Intelligence",
            "⏱️ 7-Stage Cycle & SLA Analytics",
            "🏗️ Vendor Risk & Kontraktor Renovasi",
            "🔍 Batch Evaluator & Data Explorer"
        ],
        index=0
    )

    st.markdown("---")
    st.markdown("### ⚙️ Filter Global Data")
    
    available_years = sorted(DF_GLOBAL['Tahun'].dropna().unique().tolist()) if not DF_GLOBAL.empty else [2024, 2025, 2026]
    selected_year = st.selectbox("Pilih Tahun Analisis:", ["Semua Tahun"] + [str(y) for y in available_years], index=len(available_years))
    
    if selected_year == "Semua Tahun":
        DF_ACTIVE = DF_GLOBAL.copy()
    else:
        DF_ACTIVE = DF_GLOBAL[DF_GLOBAL['Tahun'] == int(selected_year)].copy()

    st.markdown("---")
    st.markdown("""
        <div style="background: #1e293b; padding: 12px; border-radius: 8px; border: 1px solid #334155;">
            <p style="font-size: 11px; color: #94a3b8; margin: 0;">
                <b>Versi:</b> 3.3.0 (Paritas 100% Model)<br>
                <b>Basis Data:</b> 1.039 Cabang (2024–2026)<br>
                <b>SLA Target:</b> 17.0 Hari | <b>Cap Max:</b> 15 Proyek
            </p>
        </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# 6. HEADER UTAMA
# ==============================================================================
st.markdown("""
    <div class="main-header">
        <h1>🏢 PUSAT GADAI INDONESIA (PGI) — EXECUTIVE INTELLIGENCE</h1>
        <p>Sistem Analisis Terpadu Ekspansi Cabang (UPC), Optimasi Negosiasi Ruko & Smart Auto-Routing Berbasis AI</p>
    </div>
""", unsafe_allow_html=True)


# ==============================================================================
# MODUL 1: EXECUTIVE DASHBOARD
# ==============================================================================
if menu == "📊 Executive Dashboard":
    st.markdown("### 📈 Ringkasan Eksekutif & Key Performance Indicators (KPI)")
    
    tot_db = len(DF_GLOBAL)
    tot_active = len(DF_ACTIVE)
    tot_saving = DF_ACTIVE['saving_rp'].sum()
    avg_diskon = DF_ACTIVE['diskon_pct'].mean() if not DF_ACTIVE.empty else 0
    avg_durasi = DF_ACTIVE['durasi_hari'].mean() if not DF_ACTIVE.empty else 0
    
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Total Transaksi</div>
                <div class="metric-value">{tot_active:,} <span style="font-size: 14px; font-weight: 500; color: #94a3b8;">Deal</span></div>
                <div class="metric-sub">Database: {tot_db:,} Cabang</div>
            </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Total Efisiensi (Saving)</div>
                <div class="metric-value" style="color: #34d399;">{rupiah(tot_saving)}</div>
                <div class="metric-sub">Kumulatif Diskon Riil</div>
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
                <div class="metric-value" style="color: #facc15;">{avg_durasi:.1f} <span style="font-size: 14px; font-weight: 500; color: #94a3b8;">Hari</span></div>
                <div class="metric-sub">Target SLA: ≤ 17.0 Hari</div>
            </div>
        """, unsafe_allow_html=True)
    with c5:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Harmonisasi Lead Time</div>
                <div class="metric-value" style="color: #c084fc;">58.7 <span style="font-size: 14px; font-weight: 500; color: #94a3b8;">Hari</span></div>
                <div class="metric-sub">Turun dari 69 Hari (v1)</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("#### 👥 Komparasi Kinerja Tim Negosiator")
    col_bonita, col_mirza, col_lain = st.columns(3)
    
    df_bonita = DF_ACTIVE[DF_ACTIVE['negosiator'] == 'Bonita']
    df_mirza = DF_ACTIVE[DF_ACTIVE['negosiator'] == 'Mirza']
    df_lain = DF_ACTIVE[~DF_ACTIVE['negosiator'].isin(['Bonita', 'Mirza'])]
    
    with col_bonita:
        st.markdown(f"""
            <div class="metric-card" style="border-left: 4px solid #3b82f6;">
                <h4 style="color: #60a5fa; margin: 0 0 10px 0;">👩‍💼 BONITA (Senior Specialist)</h4>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Volume Deal:</span>
                    <b>{len(df_bonita)} Deal</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Total Saving:</span>
                    <b style="color: #34d399;">{rupiah(df_bonita['saving_rp'].sum())}</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Rata-rata Diskon:</span>
                    <b style="color: #38bdf8;">{df_bonita['diskon_pct'].mean():.2f}%</b>
                </div>
                <div style="display: flex; justify-content: space-between;">
                    <span style="color: #94a3b8;">Rata-rata Durasi:</span>
                    <b style="color: #facc15;">{df_bonita['durasi_hari'].mean():.1f} Hari</b>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col_mirza:
        st.markdown(f"""
            <div class="metric-card" style="border-left: 4px solid #10b981;">
                <h4 style="color: #34d399; margin: 0 0 10px 0;">👨‍💼 MIRZA (Strategic Negotiator)</h4>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Volume Deal:</span>
                    <b>{len(df_mirza)} Deal</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Total Saving:</span>
                    <b style="color: #34d399;">{rupiah(df_mirza['saving_rp'].sum())}</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Rata-rata Diskon:</span>
                    <b style="color: #38bdf8;">{df_mirza['diskon_pct'].mean():.2f}%</b>
                </div>
                <div style="display: flex; justify-content: space-between;">
                    <span style="color: #94a3b8;">Rata-rata Durasi:</span>
                    <b style="color: #facc15;">{df_mirza['durasi_hari'].mean():.1f} Hari</b>
                </div>
            </div>
        """, unsafe_allow_html=True)

    with col_lain:
        st.markdown(f"""
            <div class="metric-card" style="border-left: 4px solid #f97316;">
                <h4 style="color: #fb923c; margin: 0 0 10px 0;">👥 TIM SURVEYOR / LAPANGAN</h4>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Volume Deal:</span>
                    <b>{len(df_lain)} Deal</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Total Saving:</span>
                    <b style="color: #34d399;">{rupiah(df_lain['saving_rp'].sum())}</b>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="color: #94a3b8;">Rata-rata Diskon:</span>
                    <b style="color: #38bdf8;">{df_lain['diskon_pct'].mean():.2f}%</b>
                </div>
                <div style="display: flex; justify-content: space-between;">
                    <span style="color: #94a3b8;">Rata-rata Durasi:</span>
                    <b style="color: #facc15;">{df_lain['durasi_hari'].mean():.1f} Hari</b>
                </div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c_chart1, c_chart2 = st.columns(2)
    with c_chart1:
        st.markdown("##### 📊 Distribusi Efisiensi Diskon (%) per Negosiator")
        fig_box = px.box(
            DF_ACTIVE[DF_ACTIVE['negosiator'].isin(['Bonita', 'Mirza', 'Tim Surveyor'])],
            x='negosiator',
            y='diskon_pct',
            color='negosiator',
            color_discrete_map={'Bonita': '#3b82f6', 'Mirza': '#10b981', 'Tim Surveyor': '#f97316'},
            points="all",
            labels={'negosiator': 'Negosiator', 'diskon_pct': 'Efisiensi Diskon (%)'},
            template="plotly_dark"
        )
        fig_box.update_layout(showlegend=False, margin=dict(l=20, r=20, t=30, b=20), height=340)
        st.plotly_chart(fig_box, use_container_width=True)

    with c_chart2:
        st.markdown("##### 💰 Harga Awal Penawaran vs Harga Deal Final (Juta Rp)")
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
        fig_scat.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=340)
        st.plotly_chart(fig_scat, use_container_width=True)


# ==============================================================================
# MODUL 2: SMART AI AUTO-ROUTING & LIVE SIMULATOR (100% PARITAS MODEL)
# ==============================================================================
elif menu == "🤖 Smart AI Auto-Routing & Simulator":
    st.markdown("### 🤖 Simulator Rekomendasi & Smart Auto-Routing Negosiator Berbasis AI/MCDA")
    st.markdown("Evaluasi calon ruko baru secara objektif dengan algoritma Multi-Criteria Decision Analysis (MCDA) yang identik 100% dengan backend internal PGI.")
    
    tab_routing, tab_pred = st.tabs(["🎯 Smart Auto-Routing (MCDA Penugasan)", "📈 Prediction Engine (Target Diskon & Kemudahan)"])
    
    with tab_routing:
        col_input, col_result = st.columns([1, 1.2])
        
        with col_input:
            st.markdown("""
                <div class="glass-box">
                    <h4 style="color: #38bdf8; margin-top: 0;">📝 Parameter Calon Cabang Baru</h4>
            """, unsafe_allow_html=True)
            
            sim_wilayah = st.selectbox(
                "Wilayah / Kota Ruko:",
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
            
            st.markdown("---")
            st.markdown("<h5 style='color: #94a3b8;'>⚖️ Status Beban Kerja Aktif Tim (Max: 15)</h5>", unsafe_allow_html=True)
            
            c_b_load, c_m_load = st.columns(2)
            with c_b_load:
                load_b = st.slider("Beban Bonita Saat Ini:", 0, 15, 10, key="sld_b_load")
            with c_m_load:
                load_m = st.slider("Beban Mirza Saat Ini:", 0, 15, 10, key="sld_m_load")
                
            st.markdown("</div>", unsafe_allow_html=True)

        # Eksekusi MCDA Routing Identik
        res = execute_routing_logic(sim_wilayah, sim_harga, load_b, load_m)
        
        with col_result:
            st.markdown("#### 🎯 Hasil Keputusan & Rekomendasi Penugasan")
            
            assigned_name = res['assigned']
            badge_class = "hero-badge-bonita" if "BONITA" in assigned_name else ("hero-badge-mirza" if "MIRZA" in assigned_name else "hero-badge-tim")
            icon = "👩‍💼" if "BONITA" in assigned_name else ("👨‍💼" if "MIRZA" in assigned_name else "👥")
            
            st.markdown(f"""
                <div class="{badge_class}">
                    <div style="font-size: 13px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.9;">Rekomendasi Negosiator Terbaik</div>
                    <h2 style="margin: 6px 0 10px 0; font-size: 28px; font-weight: 800;">{icon} {assigned_name}</h2>
                    <p style="margin: 0; font-size: 13px; line-height: 1.5; opacity: 0.95;">
                        <b>Rasional AI:</b> {res['reason']}
                    </p>
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

            st.markdown("---")
            
            # Skor MCDA Perbandingan
            st.markdown(f"""
                <div style="background: #1e293b; border-radius: 10px; padding: 14px; border: 1px solid #334155;">
                    <h5 style="margin: 0 0 8px 0; color: #38bdf8;">📊 Skor Komposit MCDA (Bobot 45% Beban | 25% Wilayah | 15% Nilai | 15% SLA)</h5>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; font-size: 14px;">
                        <div>👩‍💼 <b>Skor Bonita:</b> <span style="color: #60a5fa; font-weight: 700;">{res['scores']['bonita']}</span> / 100</div>
                        <div>👨‍💼 <b>Skor Mirza:</b> <span style="color: #34d399; font-weight: 700;">{res['scores']['mirza']}</span> / 100</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            # Radar Chart
            radar_cats = ['Workload Capacity', 'Regional Track', 'Ticket Size Fit', 'SLA Velocity']
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
            fig_rad.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), showlegend=True, template="plotly_dark", margin=dict(l=30, r=30, t=20, b=20), height=250)
            st.plotly_chart(fig_rad, use_container_width=True)

    with tab_pred:
        st.markdown("#### 📈 Target Diskon Realistis & Indeks Kemudahan Wilayah")
        
        c_p_in, c_p_out = st.columns([1, 1.2])
        with c_p_in:
            p_wil = st.selectbox("Pilih Wilayah Evaluasi:", options=ALL_REGIONS, index=0, key="sb_wil_pred")
            p_harga = st.number_input("Harga Penawaran Ruko (Rp):", min_value=5_000_000, max_value=500_000_000, value=75_000_000, step=5_000_000, key="num_harga_pred")
            p_neg = st.selectbox("Pilih Persona Eksekutor:", ["General (Tim Rata-rata)", "Bonita (The High-Value Specialist)", "Mirza (The Speed Specialist)"], index=0, key="sb_neg_pred")
            
            p_neg_code = "bonita" if "Bonita" in p_neg else ("mirza" if "Mirza" in p_neg else "General")
            pred_res = execute_prediction_simulation(p_wil, p_harga, p_neg_code)

        with c_p_out:
            st.markdown(f"""
                <div class="metric-card" style="border-left: 4px solid #38bdf8;">
                    <h4 style="color: #38bdf8; margin: 0 0 8px 0;">🎯 Hasil Prediksi: {pred_res['wilayah_official']}</h4>
                    <div style="font-size: 13px; color: #94a3b8; margin-bottom: 12px;"><b>Basis Data:</b> {pred_res['basis_data']}</div>
                    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 14px;">
                        <div><b>Target Diskon Maksimal:</b> <span style="color: #34d399; font-weight: 700;">{pred_res['target_diskon_pct']:.2f}%</span></div>
                        <div><b>Potensi Efisiensi:</b> <span style="color: #34d399; font-weight: 700;">{pred_res['potensi_penghematan_fmt']}</span></div>
                        <div><b>Estimasi Durasi:</b> <b>{pred_res['est_durasi_hari']:.1f} Hari</b></div>
                        <div><b>Target Harga Net:</b> <b>{pred_res['target_harga_net_fmt']}</b></div>
                        <div><b>Skor Kemudahan:</b> <span style="color: #facc15; font-weight: 700;">{pred_res['skor_komposit']:.1f} / 100</span></div>
                        <div><b>Kategori:</b> <span style="color: #38bdf8; font-weight: 600;">{pred_res['kategori_kemudahan']}</span></div>
                    </div>
                </div>
            """, unsafe_allow_html=True)


# ==============================================================================
# MODUL 3: GEOSPATIAL & REGIONAL INTELLIGENCE
# ==============================================================================
elif menu == "🗺️ Geospatial & Regional Intelligence":
    st.markdown("### 🗺️ Analisis Spasial & Geospasial Kemudahan Negosiasi Wilayah")
    st.markdown("Pemetaan efisiensi diskon dan kecepatan durasi negosiasi di seluruh Kabupaten/Kota di Indonesia.")
    
    if not REGIONAL_STATS.empty:
        c_kategori_sum = REGIONAL_STATS['kategori'].value_counts()
        
        c_g1, c_g2, c_g3, c_g4 = st.columns(4)
        with c_g1:
            st.metric("Total Wilayah Teranalisis", f"{len(REGIONAL_STATS)} Wilayah")
        with c_g2:
            st.metric("Kategori Sangat Mudah", f"{c_kategori_sum.get('Sangat Mudah', 0)} Wilayah", "Yield Diskon Tinggi")
        with c_g3:
            st.metric("Kategori Moderat", f"{c_kategori_sum.get('Moderat', 0)} Wilayah", "SLA Standar")
        with c_g4:
            st.metric("Kategori Alot / Sulit", f"{c_kategori_sum.get('Sangat Sulit (Alot)', 0) + c_kategori_sum.get('Sulit', 0)} Wilayah", "Perlu Strategi Khusus")
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("##### 📍 Scatter Matrix: Rata-Rata Durasi (Hari) vs Rata-Rata Diskon (%)")
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
                'Moderat': '#facc15',
                'Sulit': '#f97316',
                'Sangat Sulit (Alot)': '#ef4444'
            },
            labels={'avg_durasi': 'Rata-rata Durasi Negosiasi (Hari)', 'avg_diskon': 'Rata-rata Diskon (%)'},
            template="plotly_dark"
        )
        fig_geo_scat.add_hline(y=15.0, line_dash="dash", line_color="#94a3b8", annotation_text="Benchmark Diskon 15%")
        fig_geo_scat.add_vline(x=17.0, line_dash="dash", line_color="#f87171", annotation_text="Batas SLA 17 Hari")
        fig_geo_scat.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=420)
        st.plotly_chart(fig_geo_scat, use_container_width=True)

        c_top, c_bot = st.columns(2)
        with c_top:
            st.markdown("##### 🟢 Top 10 Wilayah Paling Mudah (Highest Yield & Speed)")
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
            st.markdown("##### 🔴 Top 10 Wilayah Paling Sulit / Alot (Strict Landlord)")
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
# MODUL 4: 7-STAGE CYCLE & SLA ANALYTICS
# ==============================================================================
elif menu == "⏱️ 7-Stage Cycle & SLA Analytics":
    st.markdown("### ⏱️ Dekomposisi 7 Tahapan Siklus Waktu & Law of Diminishing Returns")
    st.markdown("Analisis komprehensif lead time sejak survei pertama hingga grand opening cabang.")
    
    stages_data = [
        {"Tahap": "1. Survei Lokasi & Verifikasi Lapangan", "Target SLA (Hari)": 5.0, "Realisasi (Hari)": 4.8, "Status": "Optimal"},
        {"Tahap": "2. Approval Komite Bisnis / Direksi", "Target SLA (Hari)": 3.0, "Realisasi (Hari)": 3.2, "Status": "Normal"},
        {"Tahap": "3. Negosiasi Sewa Ruko (Deal Final)", "Target SLA (Hari)": 17.0, "Realisasi (Hari)": 11.8, "Status": "Sangat Cepat"},
        {"Tahap": "4. Legalitas & Penandatanganan MoU", "Target SLA (Hari)": 4.0, "Realisasi (Hari)": 4.5, "Status": "Normal"},
        {"Tahap": "5. Renovasi & Fitting Kontraktor", "Target SLA (Hari)": 20.0, "Realisasi (Hari)": 24.2, "Status": "Bottleneck Terbesar"},
        {"Tahap": "6. Perekrutan & Pelatihan Karyawan", "Target SLA (Hari)": 6.0, "Realisasi (Hari)": 5.7, "Status": "Optimal"},
        {"Tahap": "7. Setup IT, Brankas & Grand Opening", "Target SLA (Hari)": 5.0, "Realisasi (Hari)": 4.5, "Status": "Optimal"},
    ]
    df_stages = pd.DataFrame(stages_data)
    
    col_st1, col_st2 = st.columns([1.2, 1])
    with col_st1:
        st.markdown("##### 📊 Perbandingan Target SLA vs Realisasi (Hari)")
        fig_bar_stage = go.Figure()
        fig_bar_stage.add_trace(go.Bar(
            y=df_stages['Tahap'],
            x=df_stages['Target SLA (Hari)'],
            name='Target SLA',
            orientation='h',
            marker_color='#3b82f6'
        ))
        fig_bar_stage.add_trace(go.Bar(
            y=df_stages['Tahap'],
            x=df_stages['Realisasi (Hari)'],
            name='Realisasi Riil',
            orientation='h',
            marker_color='#f59e0b'
        ))
        fig_bar_stage.update_layout(
            barmode='group',
            template="plotly_dark",
            margin=dict(l=20, r=20, t=30, b=20),
            height=380
        )
        st.plotly_chart(fig_bar_stage, use_container_width=True)

    with col_st2:
        st.markdown("##### 📋 Ringkasan 7 Siklus")
        st.dataframe(df_stages, use_container_width=True, hide_index=True)
        st.info("💡 **Insight Eksekutif:** Tahap 5 (Renovasi) menjadi bottleneck utama operasional dengan deviasi +4.2 hari, sedangkan Tahap 3 (Negosiasi) berkinerja unggul 5.2 hari lebih cepat dari SLA.")

    st.markdown("---")
    
    st.markdown("#### 📉 Evaluasi Law of Diminishing Returns: Durasi Negosiasi vs Efisiensi Diskon")
    
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
            fig_dim.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=320)
            st.plotly_chart(fig_dim, use_container_width=True)
            
        with c_dim2:
            st.markdown("""
                <div class="metric-card" style="height: 100%;">
                    <h4 style="color: #38bdf8; margin-top: 0;">🔬 Kesimpulan Uji Statistik</h4>
                    <p style="font-size: 13px; line-height: 1.6; color: #cbd5e1;">
                        1. <b>Sweet Spot Negosiasi:</b> Efisiensi diskon tertinggi dicapai pada rentang <b>1–10 hari</b> (Rata-rata diskon 17.8%).<br>
                        2. <b>Diminishing Returns:</b> Menambah waktu negosiasi lebih dari 14 hari tidak memberikan tambahan diskon yang signifikan (p-value > 0.05 pada uji Mann-Whitney U).<br>
                        3. <b>Rekomendasi Kebijakan:</b> Tetapkan <i>Hard Stop</i> pada hari ke-15 untuk mencegah tertundanya jadwal Grand Opening.
                    </p>
                </div>
            """, unsafe_allow_html=True)


# ==============================================================================
# MODUL 5: VENDOR RISK & KONTRAKTOR RENOVASI
# ==============================================================================
elif menu == "🏗️ Vendor Risk & Kontraktor Renovasi":
    st.markdown("### 🏗️ Analisis Vendor Risk Quadrant & Kontraktor Renovasi")
    st.markdown("Evaluasi kepatuhan biaya (cost variance) dan ketepatan waktu (lead time) kontraktor renovasi ruko.")
    
    vendor_data = [
        {"Kontraktor": "CV Cipta Karya Mandiri", "Total Proyek": 28, "Deviasi Biaya (%)": 2.1, "Keterlambatan (Hari)": 1.5, "Kategori": "Mitra Unggulan (Tier 1)"},
        {"Kontraktor": "PT Bangun Graha Utama", "Total Proyek": 24, "Deviasi Biaya (%)": 3.8, "Keterlambatan (Hari)": 2.2, "Kategori": "Mitra Unggulan (Tier 1)"},
        {"Kontraktor": "CV Sentosa Jaya Konstruksi", "Total Proyek": 19, "Deviasi Biaya (%)": 5.4, "Keterlambatan (Hari)": 4.1, "Kategori": "Kinerja Moderat (Tier 2)"},
        {"Kontraktor": "CV Prima Jaya Teknik", "Total Proyek": 15, "Deviasi Biaya (%)": 8.2, "Keterlambatan (Hari)": 6.8, "Kategori": "Risiko Tinggi (Tier 3)"},
        {"Kontraktor": "Kontraktor Trimo (Wilayah Bali)", "Total Proyek": 8, "Deviasi Biaya (%)": 18.5, "Keterlambatan (Hari)": 14.5, "Kategori": "Kasus Kritis (Red Flag)"},
    ]
    df_vendor = pd.DataFrame(vendor_data)
    
    c_v1, c_v2 = st.columns([1.2, 1])
    with c_v1:
        st.markdown("##### 📊 Kuadran Risiko Kontraktor Renovasi")
        fig_vend = px.scatter(
            df_vendor,
            x='Keterlambatan (Hari)',
            y='Deviasi Biaya (%)',
            size='Total Proyek',
            color='Kategori',
            hover_name='Kontraktor',
            color_discrete_map={
                'Mitra Unggulan (Tier 1)': '#10b981',
                'Kinerja Moderat (Tier 2)': '#3b82f6',
                'Risiko Tinggi (Tier 3)': '#f97316',
                'Kasus Kritis (Red Flag)': '#ef4444'
            },
            template="plotly_dark"
        )
        fig_vend.add_vline(x=5.0, line_dash="dash", line_color="#facc15", annotation_text="Batas Keterlambatan 5 Hari")
        fig_vend.add_hline(y=7.0, line_dash="dash", line_color="#facc15", annotation_text="Batas Deviasi Biaya 7%")
        fig_vend.update_layout(margin=dict(l=20, r=20, t=30, b=20), height=380)
        st.plotly_chart(fig_vend, use_container_width=True)

    with c_v2:
        st.markdown("##### 📋 Tabel Evaluasi Kontraktor")
        st.dataframe(df_vendor, use_container_width=True, hide_index=True)
        st.warning("⚠️ **Catatan Khusus Direksi:** Kontraktor Trimo (Bali) memerlukan evaluasi ulang kontrak akibat deviasi waktu rata-rata 14.5 hari dan pembengkakan biaya 18.5%.")


# ==============================================================================
# MODUL 6: BATCH EVALUATOR & DATA EXPLORER
# ==============================================================================
elif menu == "🔍 Batch Evaluator & Data Explorer":
    st.markdown("### 🔍 Multi-Ruko Batch Evaluator & Interactive Data Explorer")
    st.markdown("Telusuri database 1.039 cabang atau lakukan evaluasi batch ruko sekaligus.")
    
    tab_exp, tab_batch = st.tabs(["📁 Data Explorer Cabang", "📤 Batch Multi-Ruko AI Evaluator"])
    
    with tab_exp:
        st.markdown("##### 🎯 Filter & Eksplorasi Data Cabang")
        
        f_c1, f_c2, f_c3 = st.columns(3)
        with f_c1:
            filter_nego = st.multiselect("Filter Negosiator:", options=sorted(DF_ACTIVE['negosiator'].unique().tolist()), default=None)
        with f_c2:
            filter_wil = st.multiselect("Filter Wilayah:", options=sorted(DF_ACTIVE['wilayah'].unique().tolist()), default=None)
        with f_c3:
            search_query = st.text_input("Cari Cabang / No Pengajuan:", "")

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

        st.markdown(f"Menampilkan **{len(df_filtered):,}** dari total **{len(DF_ACTIVE):,}** data")
        
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
            height=400
        )
        
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            csv_buf = df_filtered[available_display_cols].to_csv(index=False).encode('utf-8')
            st.download_button("📥 Unduh Data Hasil Filter (CSV)", data=csv_buf, file_name="data_negosiasi_pgi_filtered.csv", mime="text/csv", use_container_width=True)
        with c_dl2:
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
                df_filtered[available_display_cols].to_excel(writer, index=False, sheet_name='Data_Nego')
            st.download_button("📥 Unduh Data Hasil Filter (Excel)", data=excel_buf.getvalue(), file_name="data_negosiasi_pgi_filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with tab_batch:
        st.markdown("##### 📤 Unggah Berkas untuk Evaluasi Massal (Batch AI)")
        st.markdown("Unggah berkas Excel atau CSV berisi kolom `wilayah` dan `harga_penawaran` untuk memproses rekomendasi negosiator serentak.")
        
        uploaded_file = st.file_uploader("Pilih file CSV atau Excel", type=['csv', 'xlsx', 'xls'])
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    up_df = pd.read_csv(uploaded_file)
                else:
                    up_df = pd.read_excel(uploaded_file)
                    
                st.success(f"Berhasil memuat {len(up_df)} baris data!")
                st.write("Pratinjau Data Unggahan:", up_df.head())
                
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
                            'Alasan AI': eval_res['reason'],
                            'Estimasi Diskon (%)': eval_res['est_diskon_pct'],
                            'Potensi Saving (Rp)': eval_res['potensi_penghematan_rp'],
                            'Estimasi Durasi (Hari)': eval_res['est_durasi_hari'],
                            'Status SLA': eval_res['sla_status']
                        })
                    
                    batch_res_df = pd.DataFrame(results_batch)
                    st.markdown("#### 🎉 Hasil Evaluasi Massal AI")
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
                    st.download_button("📥 Unduh Hasil Evaluasi AI (Excel)", data=b_excel_buf.getvalue(), file_name="hasil_batch_ai_routing_pgi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("Kolom wilayah atau harga penawaran tidak ditemukan. Pastikan berkas memiliki kolom 'wilayah' dan 'harga_penawaran'.")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses file: {e}")


# ==============================================================================
# FOOTER
# ==============================================================================
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #94a3b8; font-size: 12px; padding: 10px 0;">
        © 2026 <b>Pusat Gadai Indonesia (PGI)</b> — Divisi Bisnis & Ekspansi Jaringan UPC.<br>
        Dashboard Analytics & Smart Routing Engine dikembangkan oleh <b>Mukhammad Rekza Mufti</b>.
    </div>
""", unsafe_allow_html=True)
