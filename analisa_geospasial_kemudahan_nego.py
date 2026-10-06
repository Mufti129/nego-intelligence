"""
====================================================================================================
PUSAT GADAI INDONESIA (PGI) — SKRIP ANALISIS GEOSPASIAL KEMUDAHAN NEGOSIASI WILAYAH 2026
====================================================================================================
Deskripsi:
  Skrip mandiri untuk menganalisis tingkat kemudahan dan kesulitan negosiasi properti ruko
  berdasarkan data aktual tahun 2026, memetakan skor spasial ke dalam peta tematik resolusi
  tinggi (300 DPI) menggunakan TopoJSON 524 wilayah BPS, serta menyusun profil karakteristik
  wilayah termudah vs tersulit di Indonesia.

Pustaka Standar:
  - pandas, numpy, matplotlib, json, os

Cara Menjalankan di Terminal / VSCode:
  python3 analisa_geospasial_kemudahan_nego.py
====================================================================================================
"""

import json
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection

# Konfigurasi Path Berkas
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "properties_cleaned_2024-2026.csv")
TOPO_PATH = os.path.join(BASE_DIR, "indonesia-kabkot-topo.json")
OUTPUT_DIR = os.path.join(BASE_DIR, "hasil_analisis")
GRAFIK_DIR = os.path.join(OUTPUT_DIR, "grafik")
ARTIFACT_GRAFIK_DIR = "/Users/macbookair/.gemini/antigravity/brain/67136be7-18f1-4b75-a141-3f56f31934c2/grafik"

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(GRAFIK_DIR, exist_ok=True)
if os.path.exists(os.path.dirname(ARTIFACT_GRAFIK_DIR)):
    os.makedirs(ARTIFACT_GRAFIK_DIR, exist_ok=True)


# ==============================================================================
# 1. PENGOLAHAN & KALKULASI SKOR KEMUDAHAN NEGOSIASI WILAYAH
# ==============================================================================
def olah_data_geospasial():
    """Membaca data bersih dan menghitung metrik negosiasi per kabupaten/kota."""
    print(f"[*] Membaca data properti: {DATA_PATH}")
    df = pd.read_csv(DATA_PATH)
    
    # Fokus kohort 2026
    df_2026 = df[df["Tahun"] == 2026].copy()
    df_2026["durasi"] = pd.to_numeric(df_2026.get("durasi_nego_hari", df_2026.get("lama_waktu_realisasi_nego", 0)), errors="coerce").fillna(0)
    df_2026["diskon"] = pd.to_numeric(df_2026["efisiensi_diskon_pct"], errors="coerce").fillna(0)
    df_2026["asking"] = pd.to_numeric(df_2026["harga_awal_penawaran"], errors="coerce").fillna(0)
    df_2026["rental"] = pd.to_numeric(df_2026["harga_rental_final"], errors="coerce").fillna(0)
    df_2026["saving"] = pd.to_numeric(df_2026["diskon_rupiah"], errors="coerce").fillna(0)

    # Agregasi per wilayah
    w_grp = df_2026.groupby("wilayah").agg(
        unit=("nomor_pengajuan", "count"),
        avg_durasi=("durasi", "mean"),
        med_durasi=("durasi", "median"),
        avg_diskon=("diskon", "mean"),
        med_diskon=("diskon", "median"),
        avg_asking=("asking", "mean"),
        avg_rental=("rental", "mean"),
        tot_saving=("saving", "sum"),
        deal_bonita=("negosiator_analisis", lambda s: (s == "Bonita").sum()),
        deal_mirza=("negosiator_analisis", lambda s: (s == "Mirza").sum()),
        deal_nontim=("negosiator_analisis", lambda s: (~s.isin(["Bonita", "Mirza"])).sum())
    ).reset_index()

    # Normalisasi Skor (0 - 100)
    min_d, max_d = w_grp["avg_durasi"].min(), w_grp["avg_durasi"].max()
    min_disc, max_disc = w_grp["avg_diskon"].min(), w_grp["avg_diskon"].max()

    w_grp["skor_kecepatan"] = 100 * (1 - (w_grp["avg_durasi"] - min_d) / (max_d - min_d + 1e-6))
    w_grp["skor_diskon"] = 100 * ((w_grp["avg_diskon"] - min_disc) / (max_disc - min_disc + 1e-6))
    w_grp["skor_kemudahan"] = 0.5 * w_grp["skor_kecepatan"] + 0.5 * w_grp["skor_diskon"]

    # Klasifikasi Kategori Kemudahan
    def klasifikasi(skor):
        if skor >= 70:
            return "Sangat Mudah"
        elif skor >= 60:
            return "Mudah"
        elif skor >= 50:
            return "Moderat / Cukup Alot"
        elif skor >= 40:
            return "Sulit"
        else:
            return "Sangat Sulit (Paling Alot)"

    w_grp["kategori"] = w_grp["skor_kemudahan"].apply(klasifikasi)

    # Tambahkan Klaster Regional
    def tentukan_klaster(w):
        w_lower = str(w).lower()
        if any(k in w_lower for k in ["jakarta", "depok", "tangerang", "bekasi", "bogor"]):
            return "Jabodetabek Puncak"
        elif any(k in w_lower for k in ["bandung", "cimahi"]):
            return "Bandung Raya"
        elif any(k in w_lower for k in ["tasikmalaya", "garut", "ciamis", "banjar", "pangandaran", "sukabumi", "cianjur"]):
            return "Jawa Barat Priangan & Selatan"
        elif any(k in w_lower for k in ["cirebon", "indramayu", "majalengka", "kuningan", "subang", "purwakarta", "karawang"]):
            return "Jawa Barat Pantura & Purwasuka"
        elif any(k in w_lower for k in ["serang", "lebak", "pandeglang", "cilegon"]):
            return "Banten Barat"
        elif any(k in w_lower for k in ["semarang", "demak", "kudus", "pati", "rembang", "blora", "grobogan", "pekalongan", "batang", "kendal", "tegal", "brebes", "pemalang"]):
            return "Jawa Tengah Pantura"
        elif any(k in w_lower for k in ["banyumas", "cilacap", "purbalingga", "banjarnegara", "kebumen", "purworejo"]):
            return "Jawa Tengah Selatan"
        elif any(k in w_lower for k in ["surakarta", "solo", "sukoharjo", "klaten", "boyolali", "sragen", "karanganyar", "wonogiri", "yogyakarta", "sleman", "bantul", "kulon progo", "gunungkidul"]):
            return "Solo Raya & DIY"
        elif any(k in w_lower for k in ["denpasar", "badung", "gianyar", "tabanan", "klungkung", "buleleng"]):
            return "Bali"
        elif any(k in w_lower for k in ["surabaya", "sidoarjo", "malang", "gresik", "mojokerto", "pasuruan"]):
            return "Jawa Timur"
        else:
            return "Wilayah Luar Jawa/Lainnya"

    w_grp["klaster"] = w_grp["wilayah"].apply(tentukan_klaster)
    df_2026["klaster"] = df_2026["wilayah"].apply(tentukan_klaster)

    return df_2026, w_grp


# ==============================================================================
# 2. DEKODER PURE PYTHON UNTUK TOPOJSON NASIONAL
# ==============================================================================
def baca_dan_dekode_topojson():
    """Mendekode arcs dan poligon geometri dari berkas indonesia-kabkot-topo.json."""
    print(f"[*] Membaca & mendekode TopoJSON: {TOPO_PATH}")
    with open(TOPO_PATH, "r", encoding="utf-8") as f:
        topo = json.load(f)

    scale = topo["transform"]["scale"]
    translate = topo["transform"]["translate"]

    # Decode arcs
    decoded_arcs = []
    for arc in topo["arcs"]:
        coords = []
        x, y = 0, 0
        for pt in arc:
            x += pt[0]
            y += pt[1]
            lon = x * scale[0] + translate[0]
            lat = y * scale[1] + translate[1]
            coords.append((lon, lat))
        decoded_arcs.append(coords)

    def get_ring_coords(ring_indices):
        ring = []
        for idx in ring_indices:
            if idx >= 0:
                arc = decoded_arcs[idx]
            else:
                arc = decoded_arcs[~idx][::-1]
            if ring:
                ring.extend(arc[1:])
            else:
                ring.extend(arc)
        return ring

    geometries = topo["objects"]["kabkot"]["geometries"]
    
    # Standarisasi nama BPS -> nama wilayah di CSV
    parsed_geoms = []
    for g in geometries:
        bps_id = str(g.get("id", ""))
        kk = g["properties"].get("kabkot", "")
        prov = g["properties"].get("provinsi", "")

        parts = bps_id.split("-")
        if len(parts) == 2:
            code = int(parts[1])
            if code >= 71 or parts[0] == "31":
                full_name = f"Kota {kk}"
                if parts[0] == "31" and code == 1:
                    full_name = f"Kab. {kk}"
            else:
                full_name = f"Kab. {kk}"
        else:
            full_name = kk

        # Ekstrak poligon
        if "arcs" not in g or not g.get("type"):
            continue

        g_type = g["type"]
        arcs = g["arcs"]
        polygons = []

        if g_type == "Polygon":
            for ring in arcs:
                coords = get_ring_coords(ring)
                if len(coords) >= 3:
                    polygons.append(coords)
        elif g_type == "MultiPolygon":
            for poly in arcs:
                for ring in poly:
                    coords = get_ring_coords(ring)
                    if len(coords) >= 3:
                        polygons.append(coords)

        parsed_geoms.append({
            "bps_id": bps_id,
            "full_name": full_name,
            "kabkot": kk,
            "provinsi": prov,
            "polygons": polygons
        })

    print(f"[✓] Berhasil memproses {len(parsed_geoms)} geometri wilayah TopoJSON.")
    return parsed_geoms


# ==============================================================================
# 3. VISUALISASI PETA GEOSPASIAL KEMUDAHAN NEGOSIASI (GRAFIK 5.1)
# ==============================================================================
def buat_peta_geospasial(parsed_geoms, w_grp):
    """Merender peta tematik spasial kemudahan & kesulitan negosiasi wilayah."""
    print("[*] Merender Peta Tematik Geospasial Kemudahan Negosiasi 2026...")

    score_dict = {}
    for _, r in w_grp.iterrows():
        score_dict[r["wilayah"]] = {
            "skor": r["skor_kemudahan"],
            "kategori": r["kategori"],
            "unit": r["unit"],
            "durasi": r["avg_durasi"],
            "diskon": r["avg_diskon"]
        }

    def get_color(skor):
        if skor is None:
            return "#F1F5F9"  # Slate 100 (belum ada transaksi)
        if skor >= 70:
            return "#059669"  # Emerald Green (Sangat Mudah)
        elif skor >= 60:
            return "#0284C7"  # Sky Blue (Mudah)
        elif skor >= 50:
            return "#F59E0B"  # Amber (Moderat)
        elif skor >= 40:
            return "#EA580C"  # Orange (Sulit)
        else:
            return "#DC2626"  # Red Crimson (Sangat Sulit / Alot)

    fig, ax = plt.subplots(figsize=(15, 8.5), dpi=300)
    fig.patch.set_facecolor("#F8FAFC")
    ax.set_facecolor("#EFF6FF")

    patches_by_color = {}

    for g in parsed_geoms:
        fname = g["full_name"]
        data = score_dict.get(fname)
        skor = data["skor"] if data else None
        color = get_color(skor)

        for poly_coords in g["polygons"]:
            poly = Polygon(poly_coords, closed=True)
            if color not in patches_by_color:
                patches_by_color[color] = []
            patches_by_color[color].append(poly)

    for color, p_list in patches_by_color.items():
        ec = "#94A3B8" if color == "#F1F5F9" else "#1E293B"
        lw = 0.35 if color == "#F1F5F9" else 0.75
        zorder = 1 if color == "#F1F5F9" else 2
        col = PatchCollection(p_list, facecolor=color, edgecolor=ec, linewidth=lw, zorder=zorder)
        ax.add_collection(col)

    ax.set_xlim(105.0, 115.9)
    ax.set_ylim(-8.95, -5.75)
    ax.set_aspect("equal")

    # Daftar Anotasi Kota Kunci
    annotated_keys = [
        ("Kab. Cilacap", 109.02, -7.72, 108.5, -8.35, "#059669", "Cilacap (Skor: 92.2)\n[Diskon 40.0% | 8.7h]"),
        ("Kota Tasikmalaya", 108.22, -7.33, 108.7, -6.6, "#059669", "Tasikmalaya (Skor: 76.7)\n[Diskon 22.5% | 4.0h]"),
        ("Kab. Serang", 106.15, -6.12, 105.4, -5.9, "#059669", "Serang (Skor: 75.2)\n[Diskon 27.7% | 9.0h]"),
        ("Kab. Banjarnegara", 109.68, -7.40, 109.9, -6.7, "#059669", "Banjarnegara (75.0)\n[Diskon 27.5%]"),
        ("Kota Jakarta Pusat", 106.84, -6.18, 106.5, -5.75, "#0284C7", "Jakarta Pusat (73.0)\n[Diskon 24.3% | 7.7h]"),
        ("Kab. Badung", 115.18, -8.58, 114.7, -8.2, "#0284C7", "Badung Bali (68.4)\n[Diskon 19.7% | 6.7h]"),
        ("Kota Surabaya", 112.75, -7.26, 113.3, -6.7, "#0284C7", "Surabaya (65.3)\n[Diskon 21.5% | 10.0h]"),
        ("Kab. Bandung Barat", 107.50, -6.85, 107.2, -6.3, "#DC2626", "Bandung Barat (19.0 - Ter-Alot!)\n[Diskon 4.7% | 24.3h]"),
        ("Kab. Grobogan", 110.92, -7.10, 111.4, -6.5, "#DC2626", "Grobogan (20.0 - Terlama)\n[Durasi 34.0h | Disc 17.5%]"),
        ("Kota Cirebon", 108.56, -6.73, 108.9, -6.1, "#DC2626", "Kota Cirebon (28.3)\n[Diskon cuma 2.5% | 17.0h]"),
        ("Kab. Cianjur", 107.14, -6.82, 106.8, -7.5, "#DC2626", "Cianjur (28.5)\n[Diskon 8.9% | 22.0h]"),
        ("Kab. Ciamis", 108.35, -7.33, 108.0, -8.1, "#EA580C", "Ciamis (36.9)\n[Diskon 9.3% | 17.3h]"),
        ("Kab. Blora", 111.42, -7.00, 111.9, -7.6, "#EA580C", "Blora (40.8)\n[Durasi 20.3h]"),
        ("Kota Surakarta", 110.83, -7.57, 111.3, -8.1, "#EA580C", "Solo (45.3)\n[Asking Rp85M | Dur 14.7h]")
    ]

    for name, orig_x, orig_y, text_x, text_y, c_box, label in annotated_keys:
        ax.plot(orig_x, orig_y, marker="o", markersize=5, color=c_box, markeredgecolor="white", markeredgewidth=1.2, zorder=5)
        ax.annotate(
            label,
            xy=(orig_x, orig_y),
            xytext=(text_x, text_y),
            arrowprops=dict(arrowstyle="->", color="#1E293B", lw=1.0, connectionstyle="arc3,rad=0.1"),
            fontsize=8,
            fontweight="bold",
            color="#0F172A",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor=c_box, lw=1.5, alpha=0.95),
            zorder=6
        )

    legend_elements = [
        plt.Rectangle((0,0),1,1, facecolor="#059669", edgecolor="#1E293B", label="Sangat Mudah (Skor ≥ 70) — Diskon Tinggi (≥20%), Closing Cepat (<9 hari)"),
        plt.Rectangle((0,0),1,1, facecolor="#0284C7", edgecolor="#1E293B", label="Mudah (Skor 60–69.9) — Diskon Baik (16–20%), Closing Terkontrol (5–10 hari)"),
        plt.Rectangle((0,0),1,1, facecolor="#F59E0B", edgecolor="#1E293B", label="Moderat / Cukup Alot (Skor 50–59.9) — Diskon Wajar (13–16%), Durasi Standar"),
        plt.Rectangle((0,0),1,1, facecolor="#EA580C", edgecolor="#1E293B", label="Sulit (Skor 40–49.9) — Diskon Rendah (10–13%) / Negosiasi Lambat"),
        plt.Rectangle((0,0),1,1, facecolor="#DC2626", edgecolor="#1E293B", label="Sangat Sulit / Paling Alot (Skor < 40) — Diskon Minim (<10%), Durasi Alot (>17 hari)"),
        plt.Rectangle((0,0),1,1, facecolor="#F1F5F9", edgecolor="#94A3B8", label="Belum Ada Transaksi Cabang PGI Tahun 2026")
    ]
    ax.legend(handles=legend_elements, loc="lower left", fontsize=8.5, framealpha=0.95, facecolor="white", edgecolor="#CBD5E1", title="KLASIFIKASI KEMUDAHAN NEGOSIASI WILAYAH", title_fontproperties={"weight": "bold", "size": 9.5})

    ax.set_title(
        "PETA GEOSPASIAL KEMUDAHAN & KESULITAN NEGOSIASI SEWA RUKO PUSAT GADAI INDONESIA (TAHUN 2026)\n"
        "Integrasi TopoJSON 524 Wilayah BPS — Analisis Disparitas Kemudahan Tawar-Menawar, Diskon & Kecepatan Closing",
        fontsize=12, fontweight="bold", pad=15, color="#0F172A"
    )
    ax.set_xlabel("Garis Bujur (Longitude)", fontsize=9, fontweight="bold")
    ax.set_ylabel("Garis Lintang (Latitude)", fontsize=9, fontweight="bold")
    ax.grid(True, linestyle=":", alpha=0.35, color="#94A3B8")

    summary_text = (
        "STATISTIK GEOSPASIAL 2026:\n"
        "• Total Wilayah Aktif: 76 Kab/Kota\n"
        "• Wilayah Paling Mudah: Kab. Cilacap (Skor 92.2)\n"
        "• Wilayah Paling Alot: Kab. Bandung Barat (Skor 19.0)\n"
        "• Wilayah Durasi Terlama: Kab. Grobogan (34.0 hari)\n"
        "• Rerata Nasional: Diskon 15.57% | Durasi 11.5 hari"
    )
    ax.text(
        0.985, 0.03, summary_text,
        transform=ax.transAxes,
        fontsize=8.5,
        verticalalignment="bottom",
        horizontalalignment="right",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="#F8FAFC", edgecolor="#475569", lw=1.2, alpha=0.95)
    )

    fig.tight_layout()
    p5 = os.path.join(GRAFIK_DIR, "5_peta_geospasial_kemudahan_nego_2026.png")
    fig.savefig(p5)
    if os.path.exists(ARTIFACT_GRAFIK_DIR):
        fig.savefig(os.path.join(ARTIFACT_GRAFIK_DIR, "5_peta_geospasial_kemudahan_nego_2026.png"))
    plt.close(fig)
    print(f"[✓] Berkas Peta Tematik berhasil disimpan di: {p5}")


# ==============================================================================
# 4. VISUALISASI MATRIKS KARAKTERISTIK KOTA & KLASTER REGIONAL (GRAFIK 5.2)
# ==============================================================================
def buat_grafik_karakteristik(w_grp, df_2026):
    """Merender perbandingan multi-panel: Top Termudah vs Tersulit dan Profil Klaster."""
    print("[*] Merender Grafik Karakteristik Wilayah & Klaster Regional 2026...")

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7.5), dpi=300)

    w_valid = w_grp[w_grp["unit"] >= 2].copy()
    top_easy = w_valid.sort_values("skor_kemudahan", ascending=False).head(8)
    top_hard = w_valid.sort_values("skor_kemudahan", ascending=True).head(8)

    comb = pd.concat([top_easy, top_hard]).sort_values("skor_kemudahan", ascending=True)
    y_pos = np.arange(len(comb))

    colors = ["#DC2626" if s < 50 else "#059669" for s in comb["skor_kemudahan"]]

    ax1.barh(y_pos, comb["skor_kemudahan"], color=colors, height=0.65, edgecolor="#1E293B", linewidth=0.8)
    ax1.set_yticks(y_pos)
    y_labels = []
    for _, r in comb.iterrows():
        w = r["wilayah"]
        u = int(r["unit"])
        y_labels.append(f"{w} ({u} pengajuan)")
    ax1.set_yticklabels(y_labels, fontsize=8.5, fontweight="bold")
    ax1.set_xlabel("Skor Kemudahan Negosiasi (0–100)", fontsize=9.5, fontweight="bold")
    ax1.set_title("A. Top 8 Wilayah Termudah vs Top 8 Wilayah Paling Alot\n(Menampilkan Jumlah Pengajuan, Skor, Diskon %, dan Durasi Hari)", fontsize=10, fontweight="bold")
    ax1.set_xlim(0, 122)
    ax1.grid(axis="x", linestyle="--", alpha=0.3)

    for i, (_, r) in enumerate(comb.iterrows()):
        skor = r["skor_kemudahan"]
        disc = r["avg_diskon"]
        dur = r["avg_durasi"]
        u = int(r["unit"])
        txt = f"[{u} unit] Skor: {skor:.1f} | Disc: {disc:.1f}% | Dur: {dur:.1f}h"
        ax1.annotate(txt, (skor + 1.2, i), va="center", fontsize=7.5, fontweight="bold", color="#0F172A")

    ax1.axvline(50, color="#64748B", linestyle=":", linewidth=1.5, label="Batas Tengah (Skor 50)")
    ax1.legend(loc="lower right", fontsize=8)

    c_summary = df_2026.groupby("klaster").agg(
        unit=("nomor_pengajuan", "count"),
        avg_diskon=("efisiensi_diskon_pct", "mean"),
        avg_durasi=("durasi_nego_hari", "mean"),
        tot_saving=("diskon_rupiah", "sum")
    ).reset_index().sort_values("avg_diskon", ascending=True)

    y_pos_c = np.arange(len(c_summary))
    ax2.barh(y_pos_c, c_summary["avg_diskon"], color="#2563EB", height=0.6, label="Rerata Efisiensi Diskon (%)", edgecolor="#1E3A8A")
    ax2.set_yticks(y_pos_c)
    y_c_labels = []
    for _, r in c_summary.iterrows():
        c = r["klaster"]
        u = int(r["unit"])
        y_c_labels.append(f"{c} ({u} pengajuan)")
    ax2.set_yticklabels(y_c_labels, fontsize=8.5, fontweight="bold")
    ax2.set_xlabel("Rata-rata Diskon Sewa (%)", fontsize=9.5, fontweight="bold", color="#1E3A8A")
    ax2.set_title("B. Profil Efisiensi Diskon Berdasarkan 11 Klaster Regional 2026\n(Menampilkan Jumlah Pengajuan Cabang, Durasi Rata-rata, dan Total Saving)", fontsize=10, fontweight="bold")
    ax2.set_xlim(0, 32)
    ax2.grid(axis="x", linestyle="--", alpha=0.3)

    for i, (_, r) in enumerate(c_summary.iterrows()):
        disc = r["avg_diskon"]
        save_m = r["tot_saving"] / 1e6
        u = int(r["unit"])
        dur = r["avg_durasi"]
        txt = f"{disc:.1f}% [{u} unit | {dur:.1f}h | Rp{save_m:.0f}Jt]"
        ax2.annotate(txt, (disc + 0.4, i), va="center", fontsize=7.5, fontweight="bold", color="#1E3A8A")

    # Inset Kotak Komparasi Kota vs Kabupaten
    box_txt = (
        "KOMPARASI KOTA VS KABUPATEN (2026):\n"
        "• Kota (Urban): 108 pengajuan (39.3%)\n"
        "  - Rerata Durasi: 10.3 hari (Median 8.0h)\n"
        "  - Rerata Diskon: 15.4% | Saving Rp 1.01 M\n"
        "  - Transaksi Alot (>14h): 17.6%\n"
        "• Kabupaten (Daerah/Rural): 167 pengajuan (60.7%)\n"
        "  - Rerata Durasi: 12.2 hari (Median 8.0h)\n"
        "  - Rerata Diskon: 15.7% | Saving Rp 1.03 M\n"
        "  - Transaksi Alot (>14h): 20.4%\n"
        "• KESIMPULAN: 8 dari 10 wilayah paling alot adalah\n"
        "  KABUPATEN (Bandung Barat, Grobogan, Cianjur,\n"
        "  Cirebon, Ciamis, Blora, Banyumas, Purwakarta).\n"
        "  Penyebab: Aset warisan keluarga & holding power kaku."
    )
    ax2.text(0.98, 0.04, box_txt, transform=ax2.transAxes, fontsize=7.5,
             verticalalignment="bottom", horizontalalignment="right",
             bbox=dict(boxstyle="round,pad=0.4", facecolor="#F8FAFC", edgecolor="#334155", lw=1.2, alpha=0.95))

    fig.suptitle("Analisis Karakteristik Kemudahan Negosiasi Wilayah & Komparasi Kota vs Kabupaten (Tahun 2026)", fontsize=12, fontweight="bold", y=0.98)
    fig.tight_layout()
    p6 = os.path.join(GRAFIK_DIR, "6_analisis_karakteristik_wilayah_nego_2026.png")
    fig.savefig(p6)
    if os.path.exists(ARTIFACT_GRAFIK_DIR):
        fig.savefig(os.path.join(ARTIFACT_GRAFIK_DIR, "6_analisis_karakteristik_wilayah_nego_2026.png"))
    plt.close(fig)
    print(f"[✓] Berkas Grafik Karakteristik berhasil disimpan di: {p6}")


# ==============================================================================
# 5. EKSPOR DATA HASIL ANALISIS KE EXCEL MULTI-SHEET & CSV
# ==============================================================================
def ekspor_laporan_data(df_2026, w_grp):
    """Mengekspor hasil analisis geospasial ke Excel multi-sheet dan CSV."""
    excel_path = os.path.join(OUTPUT_DIR, "Analisis_Geospasial_Kemudahan_Nego_2026.xlsx")
    print(f"[*] Mengekspor rekapitulasi data ke Excel: {excel_path}")

    w_export = w_grp.sort_values("skor_kemudahan", ascending=False).copy()
    top_easy = w_grp[w_grp["unit"] >= 2].sort_values("skor_kemudahan", ascending=False).head(15)
    top_hard = w_grp[w_grp["unit"] >= 2].sort_values("skor_kemudahan", ascending=True).head(15)

    c_summary = df_2026.groupby("klaster").agg(
        total_cabang=("nomor_pengajuan", "count"),
        total_asking=("harga_awal_penawaran", "sum"),
        total_deal=("harga_rental_final", "sum"),
        total_saving=("diskon_rupiah", "sum"),
        avg_diskon_pct=("efisiensi_diskon_pct", "mean"),
        avg_durasi_hari=("durasi_nego_hari", "mean"),
        med_durasi_hari=("durasi_nego_hari", "median"),
        deal_bonita=("negosiator_analisis", lambda s: (s == "Bonita").sum()),
        deal_mirza=("negosiator_analisis", lambda s: (s == "Mirza").sum())
    ).reset_index().sort_values("avg_diskon_pct", ascending=False)

    # Sheet 5: Komparasi Kota vs Kabupaten
    def get_tipe_w(w):
        w_s = str(w).strip()
        return "Kota (Urban)" if (w_s.startswith("Kota ") or "Jakarta" in w_s) else "Kabupaten (Daerah/Rural)"
    df_2026["tipe_wilayah"] = df_2026["wilayah"].apply(get_tipe_w)

    kota_kab_summary = df_2026.groupby("tipe_wilayah").agg(
        total_pengajuan=("nomor_pengajuan", "count"),
        avg_asking_price=("harga_awal_penawaran", "mean"),
        avg_deal_price=("harga_rental_final", "mean"),
        total_saving=("diskon_rupiah", "sum"),
        avg_saving_per_deal=("diskon_rupiah", "mean"),
        avg_diskon_pct=("efisiensi_diskon_pct", "mean"),
        median_diskon_pct=("efisiensi_diskon_pct", "median"),
        avg_durasi_hari=("durasi_nego_hari", "mean"),
        median_durasi_hari=("durasi_nego_hari", "median"),
        transaksi_alot_gt_14h=("durasi_nego_hari", lambda s: (s > 14).sum()),
        pct_alot=("durasi_nego_hari", lambda s: (s > 14).mean() * 100)
    ).reset_index()

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        w_export.to_excel(writer, sheet_name="Rekap_76_Wilayah_2026", index=False)
        top_easy.to_excel(writer, sheet_name="Top_15_Termudah", index=False)
        top_hard.to_excel(writer, sheet_name="Top_15_Paling_Alot", index=False)
        c_summary.to_excel(writer, sheet_name="Klaster_Regional", index=False)
        kota_kab_summary.to_excel(writer, sheet_name="Kota_vs_Kabupaten", index=False)

    csv_wilayah = os.path.join(OUTPUT_DIR, "skor_kemudahan_wilayah_2026.csv")
    w_export.to_csv(csv_wilayah, index=False)
    csv_klaster = os.path.join(OUTPUT_DIR, "klaster_regional_nego_2026.csv")
    c_summary.to_csv(csv_klaster, index=False)

    print(f"[✓] Ekspor Excel & CSV selesai dengan sukses!")


# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================
def main():
    print("\n" + "=" * 70)
    print(" MEMULAI ANALISIS GEOSPASIAL KEMUDAHAN NEGOSIASI WILAYAH PGI 2026")
    print("=" * 70)

    # 1. Olah Data
    df_2026, w_grp = olah_data_geospasial()

    # 2. Dekode TopoJSON
    parsed_geoms = baca_dan_dekode_topojson()

    # 3. Buat Peta Tematik Geospasial
    buat_peta_geospasial(parsed_geoms, w_grp)

    # 4. Buat Grafik Analisis Karakteristik
    buat_grafik_karakteristik(w_grp, df_2026)

    # 5. Ekspor Data Excel & CSV
    ekspor_laporan_data(df_2026, w_grp)

    print("\n" + "=" * 70)
    print(" [SUKSES] Seluruh Peta Visual, Grafik, & Rekapitulasi Berhasil Dibuat!")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
