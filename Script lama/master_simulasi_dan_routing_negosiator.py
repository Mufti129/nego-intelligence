import os
import sys
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from difflib import get_close_matches

"""
========================================================================================
MASTER SYSTEM NEGOSIASI TERPADU PGI (ROUTING ENGINE + SIMULASI PERFORMA NEGOSIATOR)
========================================================================================
Deskripsi:
Script terpadu ini menggabungkan seluruh kapabilitas penugasan otomatis dan pemodelan
kinerja negosiator PGI ke dalam satu System yang interaktif dengan 7 menu utama.

Fitur & Menu Utama:
1. Smart Auto-Routing Engine (Penugasan Otomatis Ruko Baru)
2. Simulasi Prediksi Spesifik per Negosiator (Bonita / Mirza / General)
3. Analisis Head-to-Head 37 Wilayah Overlap (Bonita vs Mirza)
4. Matriks Ringkasan Performa Seluruh Tim Negosiator
5. Matriks Pemetaan Stronghold Wilayah (Bonita vs Mirza)
6. Batch Multi-Branch Evaluator (Evaluasi Beberapa Ruko Sekaligus)
7. Pengaturan Parameter Config Engine (SLA Target, Threshold High-Value, Max Load)

File Prasyarat:
'data_negosiasi_2026(Jan-Jun26).csv' (Hasil dari script cleansing_data.py)

Cara Jalankan di Terminal VSCode:
python master_simulasi_dan_routing_negosiator.py
========================================================================================
"""

DATA_FILE_DEFAULT = "data_negosiasi_2026(Jan-Jun26).csv"
DATA_FILE_FALLBACK = "data_negosiasi_2026(Jan-Jun26).csv"

# Database Centroid TopoJSON Geografis Indonesia
TOPOJSON_CENTROIDS = {
    'KOTA BANDUNG': (-6.9175, 107.6191, 'Jawa Barat'),
    'KAB. BANDUNG': (-7.0252, 107.5197, 'Jawa Barat'),
    'KAB. BANDUNG BARAT': (-6.8422, 107.4474, 'Jawa Barat'),
    'KOTA CIMAHI': (-6.8722, 107.5422, 'Jawa Barat'),
    'KOTA BEKASI': (-6.2383, 106.9756, 'Jawa Barat'),
    'KAB. BEKASI': (-6.3644, 107.1725, 'Jawa Barat'),
    'KOTA BOGOR': (-6.5971, 106.7996, 'Jawa Barat'),
    'KAB. BOGOR': (-6.5518, 106.6291, 'Jawa Barat'),
    'KOTA DEPOK': (-6.4025, 106.7942, 'Jawa Barat'),
    'KOTA CIREBON': (-6.7320, 108.5523, 'Jawa Barat'),
    'KAB. CIREBON': (-6.7589, 108.4812, 'Jawa Barat'),
    'KOTA JAKARTA SELATAN': (-6.2615, 106.8106, 'DKI Jakarta'),
    'KOTA JAKARTA PUSAT': (-6.1805, 106.8284, 'DKI Jakarta'),
    'KOTA SURABAYA': (-7.2575, 112.7521, 'Jawa Timur'),
    'KOTA MALANG': (-7.9666, 112.6326, 'Jawa Timur'),
    'KOTA KEDIRI': (-7.8167, 112.0167, 'Jawa Timur'),
    'KOTA DENPASAR': (-8.6705, 115.2126, 'Bali'),
    'KAB. BADUNG': (-8.5833, 115.1667, 'Bali')
}

# Daftar Stronghold Wilayah Historis (37 Overlap)
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

# Config Global Default (Dapat Diubah melalui Menu Config)
CONFIG = {
    'TARGET_SLA_DAYS': 17.0,
    'HIGH_VALUE_THRESHOLD': 50000000.0,
    'MAX_PIPELINE_CAPACITY': 15
}

def find_data_file():
    for path in [DATA_FILE_DEFAULT, DATA_FILE_FALLBACK]:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        p_same = os.path.join(script_dir, path)
        p_cwd = os.path.join(os.getcwd(), path)
        p_scratch = f"/workspace/scratch/{path}"
        p_knowledge = f"/workspace/knowledge/{path}"
        for p in [path, p_same, p_cwd, p_scratch, p_knowledge]:
            if os.path.exists(p):
                return p
    return None

def load_data():
    file_path = find_data_file()
    if not file_path:
        print(f" Error: File '{DATA_FILE_DEFAULT}' tidak ditemukan!")
        print(f"   Silakan jalankan 'python cleansing_data_mentah_4-v2.py' terlebih dahulu!")
        sys.exit(1)
    return pd.read_csv(file_path)

def train_ml_model(df):
    regional_stats = df.groupby('wilayah').agg(
        total_proyek=('nomor_pengajuan', 'count'),
        avg_durasi_hari=('durasi_nego_hari', 'mean'),
        max_diskon_pct=('efisiensi_diskon_pct', 'max'),
        avg_diskon_pct=('efisiensi_diskon_pct', 'mean'),
        avg_harga_awal=('harga_awal_penawaran', 'mean')
    ).reset_index()
    
    features = regional_stats[['avg_harga_awal', 'avg_diskon_pct', 'avg_durasi_hari']].values
    scaler = StandardScaler()
    features_scaled = scaler.fit_transform(features)
    
    kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
    regional_stats['cluster'] = kmeans.fit_predict(features_scaled)
    
    cluster_means = regional_stats.groupby('cluster')['avg_harga_awal'].mean().sort_values()
    ordered_clusters = cluster_means.index.tolist()
    
    cluster_mapping = {
        ordered_clusters[0]: "Cluster 2 - Fast Velocity & High Yield (Sewa Terjangkau, Nego Lincah)",
        ordered_clusters[1]: "Cluster 0 - Tight Market / High Landlord Power (Harga Sedang, Durasi Lambat)",
        ordered_clusters[2]: "Cluster 1 - Prime Commercial Urban (Sewa Tinggi, Metropolitan)"
    }
    regional_stats['cluster_name'] = regional_stats['cluster'].map(cluster_mapping)
    
    neg_regional = df.groupby(['wilayah', 'nama_negosiator_2']).agg(
        neg_cabang=('nomor_pengajuan', 'count'),
        neg_avg_durasi=('durasi_nego_hari', 'mean'),
        neg_avg_diskon=('efisiensi_diskon_pct', 'mean'),
        neg_max_diskon=('efisiensi_diskon_pct', 'max')
    ).reset_index()
    
    return regional_stats, kmeans, scaler, cluster_mapping, neg_regional

def match_region_name(user_input, all_regions):
    clean_input = user_input.strip().upper()
    clean_reg_map = {r.strip().upper(): r for r in all_regions}
    if clean_input in clean_reg_map:
        return clean_reg_map[clean_input], True
    matches = get_close_matches(clean_input, list(clean_reg_map.keys()), n=1, cutoff=0.6)
    if matches:
        return clean_reg_map[matches[0]], True
    topo_matches = get_close_matches(clean_input, list(TOPOJSON_CENTROIDS.keys()), n=1, cutoff=0.5)
    if topo_matches:
        official_topo_name = topo_matches[0]
        prov = TOPOJSON_CENTROIDS[official_topo_name][2]
        return f"{official_topo_name.title()} ({prov})", False
    return user_input.title() + " (Wilayah Baru)", False

# ------------------------------------------------------------------------------
# FITUR 1: SMART AUTO-ROUTING ENGINE
# ------------------------------------------------------------------------------
def auto_assign_engine(wilayah_input, harga_penawaran, load_bonita, load_mirza):
    clean_wil = wilayah_input.strip().upper()
    max_cap = CONFIG['MAX_PIPELINE_CAPACITY']
    high_val = CONFIG['HIGH_VALUE_THRESHOLD']
    target_sla = CONFIG['TARGET_SLA_DAYS']
    
    if clean_wil in MIRZA_STRONGHOLDS:
        if load_mirza < max_cap:
            assigned = "MIRZA"
            reason = f"Stronghold Historis Mirza di {wilayah_input} (Diskon hingga 26.25%)."
            est_diskon_pct = 18.50
            est_durasi = 23.8
        else:
            assigned = "BONITA"
            reason = f"Mirza Overloaded ({load_mirza} active). Diserahkan ke Bonita untuk menjaga SLA."
            est_diskon_pct = 16.00
            est_durasi = 24.0
            
    elif clean_wil in BONITA_STRONGHOLDS:
        if load_bonita < max_cap:
            assigned = "BONITA"
            reason = f"Stronghold Historis Bonita di {wilayah_input} (Rerata diskon 21.4%)."
            est_diskon_pct = 22.50
            est_durasi = 23.2
        else:
            assigned = "MIRZA"
            reason = f"Bonita Overloaded ({load_bonita} active). Diserahkan ke Mirza untuk pertimbangan SLA."
            est_diskon_pct = 16.50
            est_durasi = 24.5
            
    elif harga_penawaran >= high_val:
        if load_bonita < max_cap:
            assigned = "BONITA"
            reason = f"Proyek Nilai Tinggi (>= Rp {high_val:,.0f}). Bonita diprioritaskan untuk hasilkan nominal diskon Rp terbesar."
            est_diskon_pct = 19.50
            est_durasi = 23.5
        elif load_mirza < max_cap:
            assigned = "MIRZA"
            reason = f"Proyek Nilai Tinggi, namun Bonita Overloaded ({load_bonita} active). Dialihkan ke Mirza."
            est_diskon_pct = 15.00
            est_durasi = 24.2
        else:
            assigned = "TIM PENDAMPING (Dika / Salma)"
            reason = f"Kedua negosiator utama Overloaded. Rekomendasi alokasi ke tim pendamping."
            est_diskon_pct = 12.00
            est_durasi = 25.0
    else:
        if load_mirza <= load_bonita and load_mirza < max_cap:
            assigned = "MIRZA"
            reason = f"Load Balancing: Proyek standar (< Rp {high_val:,.0f}) dialokasikan ke Mirza."
            est_diskon_pct = 14.50
            est_durasi = 24.0
        elif load_bonita < max_cap:
            assigned = "BONITA"
            reason = f"Load Balancing: Dialokasikan ke Bonita (Load: {load_bonita})."
            est_diskon_pct = 18.00
            est_durasi = 23.5
        else:
            assigned = "MIRZA"
            reason = f"Kapasitas seimbang."
            est_diskon_pct = 14.00
            est_durasi = 24.0
            
    potensi_penghematan_rp = harga_penawaran * (est_diskon_pct / 100.0)
    target_harga_net = harga_penawaran - potensi_penghematan_rp
    deviasi_sla = est_durasi - target_sla
    
    return {
        'assigned': assigned,
        'reason': reason,
        'harga_penawaran': harga_penawaran,
        'est_diskon_pct': est_diskon_pct,
        'potensi_penghematan_rp': potensi_penghematan_rp,
        'target_harga_net': target_harga_net,
        'est_durasi': est_durasi,
        'deviasi_sla': deviasi_sla
    }

# ------------------------------------------------------------------------------
# FITUR 2: SIMULASI PREDIKSI NEGOSIATOR SPESIFIK
# ------------------------------------------------------------------------------
def simulate_negotiator_prediction(negosiator_choice, wilayah_input, harga_penawaran_awal, df, regional_stats, kmeans, scaler, cluster_mapping, neg_regional):
    all_regions = regional_stats['wilayah'].tolist()
    official_name, is_registered = match_region_name(wilayah_input, all_regions)
    
    reg_row = regional_stats[regional_stats['wilayah'] == official_name] if is_registered else None
    neg_match = neg_regional[
        (neg_regional['wilayah'] == official_name) & 
        (neg_regional['nama_negosiator_2'].astype(str).str.lower() == negosiator_choice.lower())
    ] if is_registered else None
    
    if neg_match is not None and len(neg_match) > 0:
        row_neg = neg_match.iloc[0]
        est_durasi = row_neg['neg_avg_durasi']
        target_diskon_pct = row_neg['neg_max_diskon']
        total_cabang_neg = row_neg['neg_cabang']
        status_data = f"Histori Langsung {negosiator_choice.title()} di {official_name} ({total_cabang_neg} Cabang)"
    elif is_registered and reg_row is not None and len(reg_row) > 0:
        row_reg = reg_row.iloc[0]
        est_durasi = row_reg['avg_durasi_hari']
        target_diskon_pct = row_reg['max_diskon_pct']
        status_data = f"Baseline Wilayah {official_name} ({negosiator_choice.title()} belum ada histori di sini)"
    else:
        avg_durasi_nat = df['durasi_nego_hari'].mean()
        avg_diskon_nat = df['efisiensi_diskon_pct'].mean()
        input_scaled = scaler.transform([[harga_penawaran_awal, avg_diskon_nat, avg_durasi_nat]])
        pred_cluster = kmeans.predict(input_scaled)[0]
        cluster_df = regional_stats[regional_stats['cluster'] == pred_cluster]
        
        target_diskon_pct = cluster_df['max_diskon_pct'].mean()
        est_durasi = cluster_df['avg_durasi_hari'].mean()
        status_data = f"Inferensi ML Cluster Wilayah Baru ({cluster_mapping.get(pred_cluster)})"
        
    if negosiator_choice.lower() == 'bonita':
        mod_diskon = target_diskon_pct * 1.05
        mod_durasi = est_durasi * 0.98
    elif negosiator_choice.lower() == 'mirza':
        mod_diskon = target_diskon_pct * 0.95
        mod_durasi = est_durasi * 1.02
    else:
        mod_diskon = target_diskon_pct
        mod_durasi = est_durasi
        
    potensi_penghematan_rp = harga_penawaran_awal * (mod_diskon / 100.0)
    target_harga_net = harga_penawaran_awal - potensi_penghematan_rp
    
    target_sla = CONFIG['TARGET_SLA_DAYS']
    sla_compliance_pct = min(100.0, (target_sla / mod_durasi) * 100.0) if mod_durasi > 0 else 100.0
    deviasi_hari = mod_durasi - target_sla
    
    subskor_kecepatan = sla_compliance_pct
    subskor_efisiensi = min(100.0, (mod_diskon / 30.0) * 100.0)
    skor_komposit = (0.5 * subskor_kecepatan) + (0.5 * subskor_efisiensi)
    
    if skor_komposit >= 80:
        kategori = "SANGAT MUDAH (Sangat Tinggi)"
    elif skor_komposit >= 70:
        kategori = "MUDAH (Tinggi)"
    elif skor_komposit >= 50:
        kategori = "MODERATE / CUKUP (Sedang)"
    else:
        kategori = "SULIT / KHUSUS (Rendah)"
        
    return {
        'negosiator': negosiator_choice.title(),
        'wilayah_official': official_name,
        'harga_penawaran_awal': harga_penawaran_awal,
        'status_data': status_data,
        'target_diskon_pct': mod_diskon,
        'potensi_penghematan_rp': potensi_penghematan_rp,
        'target_harga_net': target_harga_net,
        'est_durasi_hari': mod_durasi,
        'sla_compliance_pct': sla_compliance_pct,
        'deviasi_hari': deviasi_hari,
        'subskor_kecepatan': subskor_kecepatan,
        'subskor_efisiensi': subskor_efisiensi,
        'skor_komposit': skor_komposit,
        'kategori': kategori
    }

# ------------------------------------------------------------------------------
# FITUR 3 & 4: HEAD TO HEAD OVERLAP & RINGKASAN TIM
# ------------------------------------------------------------------------------
def show_head_to_head_overlap(df):
    print("\n" + "="*92)
    print("      ANALISIS HEAD-TO-HEAD WILAYAH OVERLAP: BONITA VS MIRZA (37 WILAYAH)")
    print("="*92)
    
    b_df = df[df['nama_negosiator_2'].astype(str).str.lower() == 'bonita']
    m_df = df[df['nama_negosiator_2'].astype(str).str.lower() == 'mirza']
    
    b_regions = set(b_df['wilayah'].unique())
    m_regions = set(m_df['wilayah'].unique())
    overlap = sorted(list(b_regions.intersection(m_regions)))
    
    print(f" Total Wilayah Overlap: {len(overlap)} Kabupaten/Kota\n")
    print(f"{'Wilayah':<24} | {'Diskon Bonita (%)':<18} | {'Diskon Mirza (%)':<17} | {'Durasi Bonita':<14} | {'Durasi Mirza'}")
    print("-"*92)
    
    b_win_diskon, m_win_diskon = 0, 0
    for reg in overlap:
        b_row = b_df[b_df['wilayah'] == reg]
        m_row = m_df[m_df['wilayah'] == reg]
        b_disc = b_row['efisiensi_diskon_pct'].mean()
        m_disc = m_row['efisiensi_diskon_pct'].mean()
        b_dur = b_row['durasi_nego_hari'].mean()
        m_dur = m_row['durasi_nego_hari'].mean()
        
        if b_disc > m_disc:
            b_win_diskon += 1
        elif m_disc > b_disc:
            m_win_diskon += 1
            
        print(f"{reg:<24} | {b_disc:>16.2f}% | {m_disc:>15.2f}% | {b_dur:>12.1f} hr | {m_dur:>10.1f} hr")
        
    print("-"*92)
    print(f"🏆 RANGKUMAN BENCHMARK DISKON OVERLAP:")
    print(f"   • Bonita menghasilkan Diskon Lebih Tinggi di : {b_win_diskon} Wilayah")
    print(f"   • Mirza menghasilkan Diskon Lebih Tinggi di  : {m_win_diskon} Wilayah")
    print("="*92 + "\n")

def show_negosiator_summary(df):
    print("\n" + "="*92)
    print("              RINGKASAN PERFORMA TIM NEGOSIATOR (PGI NEGOTIATION TEAM)")
    print("="*92)
    summary = df.groupby('nama_negosiator_2').agg(
        total_cabang=('nomor_pengajuan', 'count'),
        avg_durasi=('durasi_nego_hari', 'mean'),
        avg_diskon=('efisiensi_diskon_pct', 'mean'),
        total_penghematan=('diskon_rupiah', 'sum')
    ).reset_index().sort_values(by='total_cabang', ascending=False)
    
    print(f"{'Negosiator':<20} | {'Cabang':<8} | {'Durasi (Hr)':<12} | {'Diskon Rerata':<15} | {'Total Diskon Disetujui (Rp)':<25}")
    print("-"*92)
    for _, row in summary.iterrows():
        print(f"{row['nama_negosiator_2']:<20} | {row['total_cabang']:<8} | {row['avg_durasi']:<12.1f} | {row['avg_diskon']:<15.2f}% | Rp {row['total_penghematan']:>22,.0f}")
    print("="*92 + "\n")

# ------------------------------------------------------------------------------
# FITUR 6: BATCH MULTI-BRANCH EVALUATOR
# ------------------------------------------------------------------------------
def run_batch_evaluator(all_regions):
    print("\n" + "="*92)
    print("       BATCH MULTI-BRANCH EVALUATOR — SIMULASI SKENARIO BEBERAPA RUKO SEKALIGUS")
    print("="*92)
    print("Masukkan daftar ruko dalam format: NamaWilayah,HargaAwal (Pisahkan per ruko dengan koma)")
    print("Contoh input: Kota Bandung,75000000 | Kab. Bekasi,45000000 | Kediri,60000000\n")
    
    raw_input = input(" Masukkan Daftar Ruko (format pipe '|'): ").strip()
    if not raw_input:
        return
        
    items = raw_input.split('|')
    print("\n" + "-"*92)
    print(f"{'No':<4} | {'Wilayah':<22} | {'Harga Awal (Rp)':<18} | {'Assigned':<10} | {'Diskon (%)':<12} | {'Penghematan (Rp)':<20}")
    print("-"*92)
    
    tot_penghematan = 0.0
    tot_harga_awal = 0.0
    load_b, load_m = 10, 10
    
    for i, item in enumerate(items, 1):
        parts = item.split(',')
        if len(parts) >= 2:
            w_name = parts[0].strip()
            try:
                h_val = float(parts[1].replace('Rp', '').replace('.', '').strip())
            except:
                h_val = 50000000.0
                
            off_w, _ = match_region_name(w_name, all_regions)
            res = auto_assign_engine(off_w, h_val, load_b, load_m)
            
            if res['assigned'] == 'BONITA':
                load_b += 1
            elif res['assigned'] == 'MIRZA':
                load_m += 1
                
            tot_penghematan += res['potensi_penghematan_rp']
            tot_harga_awal += h_val
            
            print(f"{i:<4} | {off_w:<22} | Rp {h_val:>14,.0f} | {res['assigned']:<10} | {res['est_diskon_pct']:>10.2f}% | Rp {res['potensi_penghematan_rp']:>17,.0f}")
            
    print("-"*92)
    print(f" TOTAL PROYEKSI PENGHEMATAN BATCH : Rp {tot_penghematan:,.0f} (dari total penawaran Rp {tot_harga_awal:,.0f})")
    print("="*92 + "\n")

# ------------------------------------------------------------------------------
# FITUR 7: CONFIG PARAMETER MANAGER
# ------------------------------------------------------------------------------
def show_and_edit_config():
    print("\n" + "="*92)
    print("               PENGATURAN PARAMETER ENGINE (CONFIG MANAGER)")
    print("="*92)
    print(f"   1. Target SLA Days            : {CONFIG['TARGET_SLA_DAYS']:.1f} Hari")
    print(f"   2. High-Value Project Limit   : Rp {CONFIG['HIGH_VALUE_THRESHOLD']:,.0f}")
    print(f"   3. Max Pipeline Capacity      : {CONFIG['MAX_PIPELINE_CAPACITY']} Proyek Aktif")
    print("-"*92)
    print("   Pilih nomor 1-3 untuk mengubah nilai, atau ketik 'back' untuk kembali.")
    
    choice = input(" Pilihan: ").strip().lower()
    if choice == '1':
        val = input(" Masukkan Target SLA Days Baru (contoh: 17): ").strip()
        try:
            CONFIG['TARGET_SLA_DAYS'] = float(val)
            print(f"✅ Target SLA diubah ke {CONFIG['TARGET_SLA_DAYS']} Hari.\n")
        except:
            pass
    elif choice == '2':
        val = input(" Masukkan High-Value Threshold Baru (contoh: 60000000): ").strip()
        try:
            CONFIG['HIGH_VALUE_THRESHOLD'] = float(val.replace('.', ''))
            print(f"✅ Threshold High-Value diubah ke Rp {CONFIG['HIGH_VALUE_THRESHOLD']:,.0f}.\n")
        except:
            pass
    elif choice == '3':
        val = input(" Masukkan Max Pipeline Capacity Baru (contoh: 18): ").strip()
        try:
            CONFIG['MAX_PIPELINE_CAPACITY'] = int(val)
            print(f"✅ Max Pipeline Capacity diubah ke {CONFIG['MAX_PIPELINE_CAPACITY']} Proyek.\n")
        except:
            pass

# ------------------------------------------------------------------------------
# MAIN CONTROLLER CLI INTERFACE
# ------------------------------------------------------------------------------
def main():
    df = load_data()
    regional_stats, kmeans, scaler, cluster_mapping, neg_regional = train_ml_model(df)
    all_regions = regional_stats['wilayah'].tolist()
    
    print("\n========================================================================================")
    print("    🚀 MASTER SYSTEM NEGOSIASI TERPADU PGI (ROUTING ENGINE & PERFORMANCE MODEL)")
    print("========================================================================================")
    print(f"   • Database Transaksi   : {len(df)} Cabang Terdaftar")
    print(f"   • Total Wilayah Historical: {len(all_regions)} Kab/Kota")
    print(f"   • Target SLA Standard  : {CONFIG['TARGET_SLA_DAYS']:.0f} Hari (Cap-Policy)\n")
    
    while True:
        print("========================================================================================")
        print(" [MENU UTAMA MASTER NEGOSIASI SYSTEM PGI]")
        print("  1.  Smart Auto-Routing Engine (Penugasan Otomatis Ruko Baru)")
        print("  2.  Simulasi Prediksi Spesifik Negosiator (Bonita / Mirza / General)")
        print("  3.  Head-to-Head Benchmark 37 Wilayah Overlap (Bonita vs Mirza)")
        print("  4.  Matriks Ringkasan Performa Seluruh Tim Negosiator")
        print("  5.  Matriks Pemetaan Stronghold Wilayah (Bonita vs Mirza)")
        print("  6.  Batch Multi-Branch Evaluator (Simulasi Beberapa Ruko Sekaligus)")
        print("  7.  Pengaturan Config Parameter Engine")
        print("  0.  Keluar dari Program (exit)")
        print("========================================================================================")
        
        cmd = input(" Pilih Menu / Perintah: ").strip().lower()
        if cmd in ['0', 'exit', 'q', 'quit']:
            print("Terima kasih! Master System Negosiasi Selesai.")
            break
            
        elif cmd == '1':
            wil = input("\n 1. Masukkan Nama Wilayah Ruko (contoh: Kota Bandung / Kab. Bekasi / Kediri): ").strip()
            hrg = input(f" 2. Masukkan Harga Penawaran Awal Ruko di '{wil}' (contoh: 75000000): ").strip()
            try:
                hrg_num = float(hrg.replace('Rp', '').replace('.', '').replace(',', '').strip())
            except ValueError:
                print("❌ Input harga tidak valid.\n")
                continue
                
            l_bonita = input(" 3. Active Pipeline Bonita Saat Ini (Default 10): ").strip()
            l_bonita = int(l_bonita) if l_bonita.isdigit() else 10
            l_mirza = input(" 4. Active Pipeline Mirza Saat Ini (Default 10): ").strip()
            l_mirza = int(l_mirza) if l_mirza.isdigit() else 10
            
            off_w, _ = match_region_name(wil, all_regions)
            res = auto_assign_engine(off_w, hrg_num, l_bonita, l_mirza)
            
            print("\n" + "="*88)
            print(f"         RECOMMENDATION OUTPUT — REKOMENDASI PENUGASAN OTOMATIS")
            print("="*88)
            print(f"   NEGOSIATOR DIREKOMENDASIKAN : [ {res['assigned']} ]")
            print(f"   Wilayah Target               : {off_w}")
            print(f"   Harga Penawaran Awal        : Rp {res['harga_penawaran']:,.0f}")
            print(f"   Alasan Bisnis & Routing     : {res['reason']}")
            print("-"*88)
            print(f"  PROYEKSI TARGET HASIL:")
            print(f"   • Target Diskon Maksimal      : {res['est_diskon_pct']:.2f}%")
            print(f"   • Proyeksi Penghematan (Rp)  : Rp {res['potensi_penghematan_rp']:,.0f}")
            print(f"   • Target Harga Net Final     : Rp {res['target_harga_net']:,.0f}")
            print(f"   • Proyeksi Waktu Negosiasi   : {res['est_durasi']:.1f} Hari (Deviasi SLA 17 Hr: +{res['deviasi_sla']:.1f} Hr)")
            print("="*88 + "\n")
            
        elif cmd == '2':
            print("\n--- PILIH NEGOSIATOR TARGET ---")
            print(" 1. Bonita")
            print(" 2. Mirza")
            print(" 3. Tim General (Rata-Rata)")
            neg_choice = input(" Pilih Negosiator (ketik nama/angka 1-3): ").strip()
            neg_name = 'Bonita' if neg_choice == '1' else ('Mirza' if neg_choice == '2' else 'General')
            
            wil_inp = input(" Masukkan Nama Wilayah Ruko (contoh: Kota Bandung / Kediri): ").strip()
            hrg_inp = input(f" Masukkan Harga Penawaran Awal Ruko di '{wil_inp}' (contoh: 50000000): ").strip()
            try:
                hrg_num = float(hrg_inp.replace('Rp', '').replace('.', '').replace(',', '').strip())
            except ValueError:
                print("❌ Input harga tidak valid.\n")
                continue
                
            res = simulate_negotiator_prediction(neg_name, wil_inp, hrg_num, df, regional_stats, kmeans, scaler, cluster_mapping, neg_regional)
            
            status_sla_txt = "SESUAI SLA" if res['deviasi_hari'] <= 0 else f"POTENSI TERLAMBAT {res['deviasi_hari']:.1f} HARI"
            
            print("\n" + "="*88)
            print(f"      HASIL SIMULASI PREDIKSI NEGOSIASI — NEGOSIATOR: {res['negosiator'].upper()}")
            print("="*88)
            print(f"  • Penugasan Negosiator : {res['negosiator']}")
            print(f"  • Wilayah Target       : {res['wilayah_official']}")
            print(f"  • Harga Penawaran Awal : Rp {res['harga_penawaran_awal']:,.0f}")
            print(f"  • Basis Data Model     : {res['status_data']}")
            print("-"*88)
            print(f"  ESTIMASI OUTPUT & TARGET FINANSIAL:")
            print(f"   1. TARGET DISKON MAKSIMAL (%) : {res['target_diskon_pct']:.2f}%")
            print(f"      • Potensi Penghematan    : Rp {res['potensi_penghematan_rp']:,.0f}")
            print(f"      • Target Harga Net Final : Rp {res['target_harga_net']:,.0f}")
            print(f"   2. ESTIMASI SLA WAKTU NEGO   : {res['est_durasi_hari']:.1f} Hari")
            print(f"      • Status Target SLA 17 Hr : {status_sla_txt}")
            print(f"   3. SKOR KEMUDAHAN NEGO (50:50): {res['skor_komposit']:.2f} / 100 ({res['kategori']})")
            print("="*88 + "\n")
            
        elif cmd == '3':
            show_head_to_head_overlap(df)
            
        elif cmd == '4':
            show_negosiator_summary(df)
            
        elif cmd == '5':
            print("\n📌 MATRIKS PEMETAAN STRONGHOLD WILAYAH (BONITA VS MIRZA):")
            print(f"   • Mirza Strongholds ({len(MIRZA_STRONGHOLDS)} Wilayah) : {', '.join(MIRZA_STRONGHOLDS)}")
            print(f"   • Bonita Strongholds ({len(BONITA_STRONGHOLDS)} Wilayah): {', '.join(BONITA_STRONGHOLDS[:10])}, dll.\n")
            
        elif cmd == '6':
            run_batch_evaluator(all_regions)
            
        elif cmd == '7':
            show_and_edit_config()
            
        else:
            print("❌ Pilihan menu tidak valid. Pilih angka 0-7.\n")

if __name__ == '__main__':
    main()
