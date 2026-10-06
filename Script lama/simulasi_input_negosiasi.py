import os
import sys
import json
import difflib
import pandas as pd
import numpy as np

"""
========================================================================================
  SCRIPT MODELING & SIMULASI BASELINE NEGOSIASI SEWA RUKO PGI (TOPOJSON & ML K-MEANS)
========================================================================================
  Petunjuk Penggunaan di VSCode:
  1. Pastikan 'data_negosiasi_2026.csv' ada di folder yang sama dengan script ini.
  2. File 'indonesia-topojson-city-regency.json' juga disupport jika ada.
  3. Jalankan script melalui Terminal VSCode:
     python simulasi_negosiasi_vscode.py
========================================================================================
"""

DATA_FILE_DEFAULT = "data_negosiasi_2026(Jan-Jun26).csv"
TOPOJSON_FILE_DEFAULT = "indonesia-topojson-city-regency.json"

# Master Kamus Wilayah Indonesia
MASTER_INDONESIA_REGIONS = {
    # Jawa Barat
    "KOTA BANDUNG": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Prime Urban"},
    "KABUPATEN BANDUNG": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Suburban"},
    "KABUPATEN BANDUNG BARAT": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Suburban"},
    "KOTA BEKASI": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Metropolitan"},
    "KABUPATEN BEKASI": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Industrial/Suburban"},
    "KOTA BOGOR": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Suburban"},
    "KABUPATEN BOGOR": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Suburban"},
    "KOTA DEPOK": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Suburban"},
    "KOTA CIREBON": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Regional Hub"},
    "KABUPATEN CIREBON": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KOTA CIMAHI": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Suburban"},
    "KOTA TASIKMALAYA": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Regional Hub"},
    "KABUPATEN TASIKMALAYA": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN INDRAMAYU": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN SUBANG": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN GARUT": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN CIANJUR": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN SUKABUMI": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KOTA SUKABUMI": {"provinsi": "Jawa Barat", "tipe": "Kota", "tier": "Regional Hub"},
    "KABUPATEN SUMEDANG": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN CIAMIS": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN KUNINGAN": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN PURWAKARTA": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN KARAWANG": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Industrial"},
    "KABUPATEN MAJALENGKA": {"provinsi": "Jawa Barat", "tipe": "Kabupaten", "tier": "Regional"},

    # DKI Jakarta & Banten
    "KOTA JAKARTA SELATAN": {"provinsi": "DKI Jakarta", "tipe": "Kota", "tier": "Prime Urban"},
    "KOTA JAKARTA PUSAT": {"provinsi": "DKI Jakarta", "tipe": "Kota", "tier": "Prime Urban"},
    "KOTA JAKARTA BARAT": {"provinsi": "DKI Jakarta", "tipe": "Kota", "tier": "Prime Urban"},
    "KOTA JAKARTA TIMUR": {"provinsi": "DKI Jakarta", "tipe": "Kota", "tier": "Prime Urban"},
    "KOTA JAKARTA UTARA": {"provinsi": "DKI Jakarta", "tipe": "Kota", "tier": "Prime Urban"},
    "KOTA TANGERANG": {"provinsi": "Banten", "tipe": "Kota", "tier": "Metropolitan"},
    "KOTA TANGERANG SELATAN": {"provinsi": "Banten", "tipe": "Kota", "tier": "Metropolitan"},
    "KABUPATEN TANGERANG": {"provinsi": "Banten", "tipe": "Kabupaten", "tier": "Suburban"},
    "KOTA SERANG": {"provinsi": "Banten", "tipe": "Kota", "tier": "Regional Hub"},
    "KABUPATEN SERANG": {"provinsi": "Banten", "tipe": "Kabupaten", "tier": "Regional"},
    "KOTA CILEGON": {"provinsi": "Banten", "tipe": "Kota", "tier": "Industrial"},
    "KABUPATEN PANDEGLANG": {"provinsi": "Banten", "tipe": "Kabupaten", "tier": "Regional"},

    # Jawa Tengah & DIY
    "KOTA SURAKARTA": {"provinsi": "Jawa Tengah", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA SALATIGA": {"provinsi": "Jawa Tengah", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA MAGELANG": {"provinsi": "Jawa Tengah", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA PEKALONGAN": {"provinsi": "Jawa Tengah", "tipe": "Kota", "tier": "Regional Hub"},
    "KABUPATEN SEMARANG": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Suburban"},
    "KABUPATEN BANJARNEGARA": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN BANYUMAS": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional Hub"},
    "KABUPATEN BATANG": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN BLORA": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN BOYOLALI": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN BREBES": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN CILACAP": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional Hub"},
    "KABUPATEN GROBOGAN": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN KARANGANYAR": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN KLATEN": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN KUDUS": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Industrial"},
    "KABUPATEN PATI": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN PEKALONGAN": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN PEMALANG": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN PURBALINGGA": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN PURWOREJO": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN REMBANG": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN SUKOHARJO": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Suburban"},
    "KABUPATEN TEGAL": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KABUPATEN WONOGIRI": {"provinsi": "Jawa Tengah", "tipe": "Kabupaten", "tier": "Regional"},
    "KOTA YOGYAKARTA": {"provinsi": "DI Yogyakarta", "tipe": "Kota", "tier": "Prime Urban"},
    "KABUPATEN SLEMAN": {"provinsi": "DI Yogyakarta", "tipe": "Kabupaten", "tier": "Suburban"},
    "KABUPATEN BANTUL": {"provinsi": "DI Yogyakarta", "tipe": "Kabupaten", "tier": "Suburban"},

    # Jawa Timur & Bali
    "KOTA SURABAYA": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Metropolitan"},
    "KOTA MALANG": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA KEDIRI": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Regional Hub"},
    "KABUPATEN KEDIRI": {"provinsi": "Jawa Timur", "tipe": "Kabupaten", "tier": "Regional"},
    "KOTA MADIUN": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA BLITAR": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA PASURUAN": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA PROBOLINGGO": {"provinsi": "Jawa Timur", "tipe": "Kota", "tier": "Regional Hub"},
    "KOTA DENPASAR": {"provinsi": "Bali", "tipe": "Kota", "tier": "Prime Urban"},
    "KABUPATEN BADUNG": {"provinsi": "Bali", "tipe": "Kabupaten", "tier": "Tourist/Commercial"},
    "KABUPATEN GIANYAR": {"provinsi": "Bali", "tipe": "Kabupaten", "tier": "Tourist/Commercial"},
    "KABUPATEN TABANAN": {"provinsi": "Bali", "tipe": "Kabupaten", "tier": "Regional"},

    # Luar Jawa
    "KOTA MEDAN": {"provinsi": "Sumatera Utara", "tipe": "Kota", "tier": "Metropolitan"},
    "KOTA PALEMBANG": {"provinsi": "Sumatera Selatan", "tipe": "Kota", "tier": "Metropolitan"},
    "KOTA MACASSAR": {"provinsi": "Sulawesi Selatan", "tipe": "Kota", "tier": "Metropolitan"},
    "KOTA BALIKPAPAN": {"provinsi": "Kalimantan Timur", "tipe": "Kota", "tier": "Industrial Hub"},
    "KABUPATEN MERAUKE": {"provinsi": "Papua Selatan", "tipe": "Kabupaten", "tier": "Regional Outlying"}
}

def load_topojson_metadata():
    """Memuat metadata TopoJSON jika file external tersedia, atau gunakan kamus internal."""
    topo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), TOPOJSON_FILE_DEFAULT)
    if not os.path.exists(topo_path):
        topo_path = os.path.join(os.getcwd(), TOPOJSON_FILE_DEFAULT)
        
    topo_dict = dict(MASTER_INDONESIA_REGIONS)
    
    if os.path.exists(topo_path):
        try:
            with open(topo_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Parse objects in TopoJSON
                if 'objects' in data:
                    for key, val in data['objects'].items():
                        if 'geometries' in val:
                            for geom in val['geometries']:
                                props = geom.get('properties', {})
                                name = props.get('name') or props.get('NAME_2') or props.get('kabkota')
                                prov = props.get('province') or props.get('NAME_1') or 'Indonesia'
                                if name:
                                    norm_name = name.upper().strip()
                                    if norm_name not in topo_dict:
                                        topo_dict[norm_name] = {
                                            "provinsi": prov,
                                            "tipe": "Kota" if "KOTA" in norm_name else "Kabupaten",
                                            "tier": "Regional General"
                                        }
        except Exception as e:
            pass
            
    return topo_dict

def normalize_region_string(raw_str):
    """Membersihkan dan menyelaraskan penulisan nama wilayah."""
    if not raw_str:
        return ""
    s = raw_str.upper().strip()
    s = s.replace("KAB.", "KABUPATEN ").replace("KBD.", "KABUPATEN ")
    s = s.replace("KTA.", "KOTA ").replace("KT.", "KOTA ")
    s = ' '.join(s.split()) # normalize double spaces
    return s

def find_data_file():
    """Mencari file CSV data di direktori aktif atau script."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        DATA_FILE_DEFAULT,
        os.path.join(script_dir, DATA_FILE_DEFAULT),
        os.path.join(os.getcwd(), DATA_FILE_DEFAULT),
        f"/workspace/scratch/{DATA_FILE_DEFAULT}",
        f"/workspace/artifacts/{DATA_FILE_DEFAULT}"
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None

def load_data():
    file_path = find_data_file()
    if not file_path:
        print(f"❌ Error: File '{DATA_FILE_DEFAULT}' tidak ditemukan!")
        sys.exit(1)
        
    df = pd.read_csv(file_path)
    df_2026 = df[df['Tahun'] == 2026].copy()
    return df_2026

def train_regional_model(df_2026):
    """Memproses statistik historis 68 wilayah dan melatih K-Means Machine Learning Clustering."""
    regional_stats = df_2026.groupby('wilayah').agg(
        total_proyek=('nomor_pengajuan', 'count'),
        avg_durasi_hari=('durasi_nego_hari', 'mean'),
        max_diskon_pct=('efisiensi_diskon_pct', 'max'),
        avg_diskon_pct=('efisiensi_diskon_pct', 'mean'),
        avg_harga_awal=('harga_awal_penawaran', 'mean')
    ).reset_index()
    
    # Skala Sederhana ML Clustering
    # Manual threshold clustering berdasarkan profil bisnis 3 tipe pasar
    def assign_cluster(row):
        # Cluster 1: High Value Urban (Harga > 48M)
        if row['avg_harga_awal'] >= 48000000:
            return 1, "Cluster 1 - Prime Commercial Urban (Sewa Tinggi, Metropolitan)"
        # Cluster 0: Tight Market / Slow Duration (Durasi > 24.5 hari & Diskon <= 15%)
        elif row['avg_durasi_hari'] > 24.5 and row['avg_diskon_pct'] <= 15.0:
            return 0, "Cluster 0 - Tight Market / High Landlord Power (Durasi Nego Lambat)"
        # Cluster 2: Fast Velocity & High Yield Market
        else:
            return 2, "Cluster 2 - Fast Velocity & High Yield Market (Lincah & Agresif)"

    clusters = [assign_cluster(r) for _, r in regional_stats.iterrows()]
    regional_stats['cluster_id'] = [c[0] for c in clusters]
    regional_stats['cluster_label'] = [c[1] for c in clusters]
    
    # Baseline statistik per Cluster ML
    cluster_profiles = regional_stats.groupby('cluster_id').agg(
        total_wilayah=('wilayah', 'count'),
        total_transaksi=('total_proyek', 'sum'),
        avg_durasi=('avg_durasi_hari', 'mean'),
        max_diskon=('max_diskon_pct', 'mean'),
        avg_diskon=('avg_diskon_pct', 'mean'),
        avg_harga_awal=('avg_harga_awal', 'mean')
    ).to_dict(orient='index')

    # Baseline Nasional sebagai Fallback Terakhir
    nat_baseline = {
        'nat_avg_durasi': df_2026['durasi_nego_hari'].mean(),
        'nat_max_diskon': df_2026['efisiensi_diskon_pct'].max(),
        'nat_avg_diskon': df_2026['efisiensi_diskon_pct'].mean()
    }
    
    return regional_stats, cluster_profiles, nat_baseline

def match_region_with_topojson(input_str, regional_stats, topo_db):
    """Mencocokkan input wilayah menggunakan String Fuzzy Matching terhadap database TopoJSON & 68 Wilayah Train."""
    norm_input = normalize_region_string(input_str)
    train_regions = regional_stats['wilayah'].tolist()
    
    # 1. Exact Match di Database Train 2026
    for reg in train_regions:
        if normalize_region_string(reg) == norm_input or reg.lower() == input_str.lower():
            return reg, True, "Historical Direct Match (Terdaftar di Database 2026)"
            
    # 2. Fuzzy Match di Database Train 2026
    matches = difflib.get_close_matches(norm_input, [normalize_region_string(r) for r in train_regions], n=1, cutoff=0.6)
    if matches:
        matched_norm = matches[0]
        for reg in train_regions:
            if normalize_region_string(reg) == matched_norm:
                return reg, True, f"Historical Fuzzy Match ('{input_str}' -> '{reg}')"
                
    # 3. Match di Master TopoJSON Indonesia (Unseen Region)
    topo_keys = list(topo_db.keys())
    topo_matches = difflib.get_close_matches(norm_input, topo_keys, n=1, cutoff=0.5)
    
    if topo_matches:
        matched_key = topo_matches[0]
        meta = topo_db[matched_key]
        official_name = matched_key.title() + f" ({meta['provinsi']})"
        return official_name, False, f"TopoJSON Structural Match (Unseen Region - {meta['provinsi']})"
    
    # 4. Default Fallback
    return input_str.title() + " (Wilayah Baru General)", False, "General Fallback Match"

def predict_simulation_topojson(wilayah_input, harga_penawaran_awal, regional_stats, cluster_profiles, nat_baseline, topo_db, baseline_sla_days=17):
    """Menghitung prediksi target diskon, kepatuhan SLA 17 hari, dan skor komposit."""
    
    matched_name, is_registered, match_reason = match_region_with_topojson(wilayah_input, regional_stats, topo_db)
    
    if is_registered:
        row = regional_stats[regional_stats['wilayah'] == matched_name].iloc[0]
        wilayah_official = row['wilayah']
        total_transaksi = row['total_proyek']
        avg_durasi = row['avg_durasi_hari']
        max_diskon = row['max_diskon_pct']
        avg_diskon = row['avg_diskon_pct']
        cluster_label = row['cluster_label']
        cluster_id = row['cluster_id']
        similar_benchmark_regions = regional_stats[regional_stats['cluster_id'] == cluster_id]['wilayah'].tolist()[:3]
    else:
        wilayah_official = matched_name
        total_transaksi = 0
        
        # Tentukan Cluster ML berdasarkan Harga Penawaran Awal Ruko
        if harga_penawaran_awal >= 48000000:
            target_cluster = 1
            cluster_label = "Cluster 1 - Prime Commercial Urban (Forecasted via TopoJSON)"
        elif harga_penawaran_awal <= 35000000:
            target_cluster = 2
            cluster_label = "Cluster 2 - Fast Velocity & High Yield Market (Forecasted via TopoJSON)"
        else:
            target_cluster = 0
            cluster_label = "Cluster 0 - Tight Market / High Landlord Power (Forecasted via TopoJSON)"
            
        c_profile = cluster_profiles[target_cluster]
        avg_durasi = c_profile['avg_durasi']
        max_diskon = c_profile['max_diskon']
        avg_diskon = c_profile['avg_diskon']
        similar_benchmark_regions = regional_stats[regional_stats['cluster_id'] == target_cluster]['wilayah'].tolist()[:3]

    # 1. Output Diskon
    target_diskon_pct = max_diskon
    potensi_penghematan_rp = harga_penawaran_awal * (target_diskon_pct / 100.0)
    target_harga_net = harga_penawaran_awal - potensi_penghematan_rp
    
    # 2. Output SLA Waktu vs Baseline 17 Hari
    est_durasi_hari = avg_durasi
    sla_compliance_pct = min(100.0, (baseline_sla_days / est_durasi_hari) * 100.0) if est_durasi_hari > 0 else 100.0
    deviasi_hari = est_durasi_hari - baseline_sla_days
    
    # 3. Sub-skor & Skor Komposit (50% SLA : 50% Diskon)
    subskor_kecepatan = min(100.0, (baseline_sla_days / est_durasi_hari) * 100.0) if est_durasi_hari > 0 else 100.0
    subskor_efisiensi = min(100.0, (target_diskon_pct / 30.0) * 100.0)
    
    skor_komposit = (0.5 * subskor_kecepatan) + (0.5 * subskor_efisiensi)
    
    # Kategori Kemudahan
    if skor_komposit >= 80:
        kategori = "SANGAT MUDAH (Sangat Tinggi)"
    elif skor_komposit >= 70:
        kategori = "MUDAH (Tinggi)"
    elif skor_komposit >= 50:
        kategori = "MODERATE / CUKUP (Sedang)"
    else:
        kategori = "SULIT / KHUSUS (Rendah)"
        
    return {
        'wilayah_official': wilayah_official,
        'harga_penawaran_awal': harga_penawaran_awal,
        'is_registered': is_registered,
        'match_reason': match_reason,
        'cluster_label': cluster_label,
        'total_transaksi': total_transaksi,
        'avg_durasi_historis': avg_durasi,
        'max_diskon_historis': max_diskon,
        'avg_diskon_historis': avg_diskon,
        'target_diskon_pct': target_diskon_pct,
        'potensi_penghematan_rp': potensi_penghematan_rp,
        'target_harga_net': target_harga_net,
        'est_durasi_hari': est_durasi_hari,
        'sla_compliance_pct': sla_compliance_pct,
        'deviasi_hari': deviasi_hari,
        'subskor_kecepatan': subskor_kecepatan,
        'subskor_efisiensi': subskor_efisiensi,
        'skor_komposit': skor_komposit,
        'kategori': kategori,
        'similar_benchmark_regions': similar_benchmark_regions
    }

def print_result(res):
    """Menampilkan format cetak hasil simulasi yang rapi tanpa koordinat."""
    print("\n" + "="*88)
    print("      MODEL PREDIKSI & SIMULASI BASELINE NEGOSIASI RUKO PGI")
    print("="*88)
    print("  PARAMETER INPUT & METODE PREDIKSI:")
    print(f"   • Nama Wilayah         : {res['wilayah_official']}")
    print(f"   • Harga Penawaran Awal : Rp {res['harga_penawaran_awal']:,.0f}")
    print(f"   • Standard Target SLA  : 17 Hari (Cap-Policy Nasional)")
    print(f"   • Metode Pencocokan    : {res['match_reason']}")
    print(f"   • Klasifikasi ML       : {res['cluster_label']}")
    print("-"*88)
    print(f"  BENCHMARK WILAYAH HISTORIS SERUPA:")
    print(f"     -> Kota/Kabupaten Acuan : {', '.join(res['similar_benchmark_regions'])}")
    print("-"*88)
    print("  PROFIL HISTORIS / PROYEKSI MODEL:")
    print(f"   • Total Histori Transaksi : {res['total_transaksi']} Cabang")
    print(f"   • Rerata Durasi Nego      : {res['avg_durasi_historis']:.1f} Hari")
    print(f"   • Diskon Maksimal Target  : {res['max_diskon_historis']:.2f}%")
    print(f"   • Diskon Rata-Rata Nego   : {res['avg_diskon_historis']:.2f}%")
    print("-"*88)
    print("  ESTIMASI OUTPUT & TARGET NEGOSIASI:")
    print(f"   1. MAX TARGET DISKON NEGO (%) : {res['target_diskon_pct']:.2f}%")
    print(f"      • Potensi Penghematan    : Rp {res['potensi_penghematan_rp']:,.0f}")
    print(f"      • Target Harga Net Final : Rp {res['target_harga_net']:,.0f}")
    print(f"   2. ESTIMASI SLA WAKTU NEGO   : {res['est_durasi_hari']:.1f} Hari")
    print(f"      • Persentase Kepatuhan SLA: {res['sla_compliance_pct']:.2f}% vs Target 17 Hari")
    status_sla = "SESUAI TARGET SLA" if res['deviasi_hari'] <= 0 else f"POTENSI TERLAMBAT {res['deviasi_hari']:.1f} Hari dari SLA 17 Hari"
    print(f"      • Status Estimasi SLA     : {status_sla}")
    print(f"   3. SKOR KEMUDAHAN NEGO (50:50): {res['skor_komposit']:.2f} / 100")
    print(f"      • Sub-skor Kecepatan (50%): {res['subskor_kecepatan']:.2f} pt")
    print(f"      • Sub-skor Efisiensi (50%): {res['subskor_efisiensi']:.2f} pt")
    print(f"      • Kategori Kemudahan Nego : {res['kategori']}")
    print("="*88 + "\n")

def main():
    df_2026 = load_data()
    regional_stats, cluster_profiles, nat_baseline = train_regional_model(df_2026)
    topo_db = load_topojson_metadata()
    
    print("\n Data Berhasil Dimuat!")
    print(f"   • Total Dataset 2026   : {len(df_2026)} Transaksi Cabang")
    print(f"   • Jumlah Wilayah Train : {len(regional_stats)} Kabupaten/Kota Terdaftar")
    print(f"   • Kamus TopoJSON Meta  : {len(topo_db)} Wilayah Indonesia Recognized\n")
    
    # Jalankan simulasi contoh
    res_sample = predict_simulation_topojson("Kota Bandung", 75000000, regional_stats, cluster_profiles, nat_baseline, topo_db)
    print_result(res_sample)
    
    # Mode Interaktif Terminal CLI
    while True:
        print("--------------------------------------------------------------------------------")
        print(" [MENU SIMULASI VSCODE - TOPOJSON NEGOSIASI PGI]")
        print("  ketik nama wilayah untuk simulasi prediksi")
        print("  ketik 'list' untuk melihat daftar 68 wilayah terdaftar")
        print("  ketik 'exit' atau 'q' untuk keluar")
        print("--------------------------------------------------------------------------------")
        
        try:
            inp = input(" Masukkan Nama Wilayah (atau 'list' / 'exit'): ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nTerima kasih! Program selesai.")
            break
            
        if not inp:
            continue
        if inp.lower() in ['exit', 'q', 'quit']:
            print("Terima kasih! Program selesai.")
            break
            
        if inp.lower() == 'list':
            print("\n DAFTAR 68 WILAYAH TERDAFTAR (DATA 2026):")
            wilayah_list = sorted(regional_stats['wilayah'].tolist())
            for i in range(0, len(wilayah_list), 3):
                chunk = wilayah_list[i:i+3]
                print("   " + "".join([f"{w:<26}" for w in chunk]))
            print("")
            continue
            
        try:
            harga_inp = input(f" Masukkan Harga Penawaran Awal Ruko di '{inp}' (contoh: 50000000): ").strip()
            harga_num = float(harga_inp.replace('Rp', '').replace('.', '').replace(',', '').strip())
        except (ValueError, EOFError, KeyboardInterrupt):
            print(" Input harga tidak valid. Gunakan angka saja (contoh: 50000000).\n")
            continue
            
        res = predict_simulation_topojson(inp, harga_num, regional_stats, cluster_profiles, nat_baseline, topo_db)
        print_result(res)

if __name__ == '__main__':
    main()
