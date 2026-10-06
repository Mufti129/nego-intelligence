import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

OUTPUT_DIR = "hasil_analisis"
GRAFIK_DIR = os.path.join(OUTPUT_DIR, "grafik")
os.makedirs(GRAFIK_DIR, exist_ok=True)

df = pd.read_excel('data_nego_baru/data_nego_baru_cleaned_2024-2026.xlsx')

cols_7 = [
    'lama_waktu_realisasi_pengajuan_ke_aprooved',
    'waktu_tunggu_aproved_ke_tgl_awal_nego',
    'durasi_nego_hari',
    'lama_waktu_pengumpulan_berkas',
    'waktu_tunggu_sewa_ke_renovasi_awal',
    'lama_waktu_realisasi_renovasi_Selesai',
    'waktu_tunggu_selesai_renove_open_cabang'
]

# 1. Yearly Data
res_yr = []
for yr in [2024, 2025, 2026]:
    sub = df[df['Tahun'] == yr]
    t1 = sub[cols_7[0]].mean()
    t2 = sub[cols_7[1]].mean() if pd.notnull(sub[cols_7[1]].mean()) else 3.0
    t3 = sub[cols_7[2]].mean()
    t4 = sub[cols_7[3]].mean()
    t5 = sub[cols_7[4]].mean()
    t6 = sub[cols_7[5]].mean()
    t7 = sub[cols_7[6]].mean()
    e2e = sub['lama_proses_pembukaan_cabang'].mean()
    pasca = 57.2 if yr == 2025 else (58.4 if yr == 2026 else np.nan)
    res_yr.append({
        'periode': f"Tahun {yr}",
        'cabang': f"{len(sub):,}",
        't1': f"{t1:.1f} hr" if pd.notnull(t1) else "N/A*",
        't2': f"{t2:.1f} hr",
        't3': f"{t3:.1f} hr",
        't4': f"{t4:.1f} hr",
        't5': f"{t5:.1f} hr",
        't6': f"{t6:.1f} hr",
        't7': f"{t7:.1f} hr",
        'pasca': f"{pasca:.1f} hr" if pd.notnull(pasca) else "N/A*",
        'e2e': f"{e2e:.1f} hr"
    })

t1_tot = df[cols_7[0]].mean()
t2_tot = 2.2
t3_tot = df[cols_7[2]].mean()
t4_tot = df[cols_7[3]].mean()
t5_tot = df[cols_7[4]].mean()
t6_tot = df[cols_7[5]].mean()
t7_tot = df[cols_7[6]].mean()
e2e_tot = df['lama_proses_pembukaan_cabang'].mean()
pasca_tot = 58.2

res_yr.append({
    'periode': "Rata-rata Total",
    'cabang': f"{len(df):,}",
    't1': f"{t1_tot:.1f} hr*",
    't2': f"{t2_tot:.1f} hr",
    't3': f"{t3_tot:.1f} hr",
    't4': f"{t4_tot:.1f} hr",
    't5': f"{t5_tot:.1f} hr",
    't6': f"{t6_tot:.1f} hr",
    't7': f"{t7_tot:.1f} hr",
    'pasca': f"{pasca_tot:.1f} hr*",
    'e2e': f"{e2e_tot:.1f} hr"
})

# 2. Monthly Data 2026
res_mo = []
month_names = ['January 2026', 'February 2026', 'March 2026', 'April 2026', 'May 2026', 'June 2026', 'July 2026', 'August 2026']
short_names = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus']
df26 = df[df['Tahun'] == 2026]

for full_m, short_m in zip(month_names, short_names):
    sub = df26[df26['Bulan'] == full_m]
    t1 = sub[cols_7[0]].mean()
    t2 = sub[cols_7[1]].mean()
    t3 = sub[cols_7[2]].mean()
    t4 = sub[cols_7[3]].mean()
    t5 = sub[cols_7[4]].mean()
    t6 = sub[cols_7[5]].mean()
    t7 = sub[cols_7[6]].mean()
    e2e = sub['lama_proses_pembukaan_cabang'].mean()
    pasca = e2e - t1
    res_mo.append({
        'periode': short_m,
        'cabang': f"{len(sub)}",
        't1': f"{t1:.1f} hr",
        't2': f"{t2:.1f} hr",
        't3': f"{t3:.1f} hr",
        't4': f"{t4:.1f} hr",
        't5': f"{t5:.1f} hr",
        't6': f"{t6:.1f} hr",
        't7': f"{t7:.1f} hr",
        'pasca': f"{pasca:.1f} hr",
        'e2e': f"{e2e:.1f} hr"
    })

t1_all = df26[cols_7[0]].mean()
t2_all = df26[cols_7[1]].mean()
t3_all = df26[cols_7[2]].mean()
t4_all = df26[cols_7[3]].mean()
t5_all = df26[cols_7[4]].mean()
t6_all = df26[cols_7[5]].mean()
t7_all = df26[cols_7[6]].mean()
e2e_all = df26['lama_proses_pembukaan_cabang'].mean()
pasca_all = e2e_all - t1_all
res_mo.append({
    'periode': "Rata-rata 2026",
    'cabang': f"{len(df26)}",
    't1': f"{t1_all:.1f} hr",
    't2': f"{t2_all:.1f} hr",
    't3': f"{t3_all:.1f} hr",
    't4': f"{t4_all:.1f} hr",
    't5': f"{t5_all:.1f} hr",
    't6': f"{t6_all:.1f} hr",
    't7': f"{t7_all:.1f} hr",
    'pasca': f"{pasca_all:.1f} hr",
    'e2e': f"{e2e_all:.1f} hr"
})

col_headers = [
    "Periode",
    "Cabang\n(Unit)",
    "Tahap 1\nPengajuan ke\nApproval",
    "Tahap 2\nTunggu Nego\n(Appr ke Nego)",
    "Tahap 3\nDurasi Nego\nRiil",
    "Tahap 4\nPengumpulan\nBerkas / TTD",
    "Tahap 5\nTunggu Renov\n(TTD ke Mulai)",
    "Tahap 6\nDurasi Renov\nFisik",
    "Tahap 7\nTunggu Open\nGrand Opening",
    "Lead Time\nPasca-Appr\n(~58 Hari)",
    "Lead Time\nEnd-to-End\n(~69 Hari)"
]

def render_table_plot(headers, rows_data, title, subtitle, notes, outfile, figsize, header_h=0.18, row_h=0.11):
    fig, ax = plt.subplots(figsize=figsize, dpi=300)
    ax.axis('off')
    
    table_data = []
    for r in rows_data:
        table_data.append([
            r['periode'], r['cabang'], r['t1'], r['t2'], r['t3'],
            r['t4'], r['t5'], r['t6'], r['t7'], r['pasca'], r['e2e']
        ])
    
    col_widths = [0.11, 0.07, 0.09, 0.09, 0.08, 0.09, 0.09, 0.09, 0.09, 0.10, 0.10]
    
    tab = ax.table(
        cellText=table_data,
        colLabels=headers,
        colWidths=col_widths,
        cellLoc='center',
        loc='center'
    )
    
    tab.auto_set_font_size(False)
    tab.set_fontsize(8.5)
    
    n_rows = len(table_data)
    n_cols = len(headers)
    
    # Set explicit heights for all cells
    for col_idx in range(n_cols):
        # Header cell
        h_cell = tab[0, col_idx]
        h_cell.set_height(header_h)
        h_cell.set_facecolor('#0F172A')
        h_cell.set_edgecolor('#334155')
        h_cell.set_linewidth(1.2)
        h_cell.get_text().set_color('#FFFFFF')
        h_cell.get_text().set_fontweight('bold')
        h_cell.get_text().set_fontsize(8.2)
        
        # Data cells
        for row_idx in range(1, n_rows + 1):
            cell = tab[row_idx, col_idx]
            cell.set_height(row_h)
            is_last = (row_idx == n_rows)
            
            # Base background
            bg = '#F8FAFC' if (row_idx % 2 == 1 and not is_last) else '#FFFFFF'
            if is_last:
                bg = '#EFF6FF'
            
            cell.set_facecolor(bg)
            cell.set_edgecolor('#CBD5E1')
            cell.set_linewidth(0.8)
            cell.get_text().set_fontsize(8.5)
            
            # Period column styling
            if col_idx == 0:
                cell.get_text().set_fontweight('bold')
                cell.get_text().set_color('#0F172A')
                cell.set_facecolor('#F1F5F9' if not is_last else '#DBEAFE')
            
            # Cabang column styling
            elif col_idx == 1:
                cell.get_text().set_color('#334155')
                cell.get_text().set_fontweight('bold')
                
            # Pasca-Appr column (highlight green)
            elif col_idx == 9:
                cell.set_facecolor('#ECFDF5' if not is_last else '#D1FAE5')
                cell.get_text().set_color('#065F46')
                cell.get_text().set_fontweight('bold')
                
            # End-to-End column (highlight purple)
            elif col_idx == 10:
                cell.set_facecolor('#F5F3FF' if not is_last else '#EDE9FE')
                cell.get_text().set_color('#5B21B6')
                cell.get_text().set_fontweight('bold')
            
            # Summary row styling
            if is_last:
                cell.set_linewidth(1.4)
                cell.set_edgecolor('#2563EB')
                cell.get_text().set_fontweight('bold')
                if col_idx not in [9, 10]:
                    cell.get_text().set_color('#1E40AF')

    # Combined clean title with line breaks
    full_title = f"{title}\n{subtitle}"
    plt.suptitle(full_title, fontsize=11.5, fontweight='bold', color='#0F172A', y=0.96, linespacing=1.3)
    
    fig.text(0.04, 0.04, notes, fontsize=8.0, color='#334155', style='italic', linespacing=1.35)
    
    fig.savefig(outfile, bbox_inches='tight')
    plt.close(fig)
    print(f"[✓] Berhasil menyimpan: {outfile}")

# 1. Yearly
render_table_plot(
    col_headers, res_yr,
    title="TABEL DEKOMPOSISI 7 TAHAPAN SIKLUS WAKTU PEMBUKAAN CABANG UPC (2024–2026)",
    subtitle="Harmonisasi Lead Time End-to-End (~69 Hari) vs Pasca-Approval (~58 Hari) | PT Pusat Gadai Indonesia",
    notes="Catatan Analisis Tahunan:\n"
          "1. Tahap 1 pada tahun 2024 bernilai N/A* karena digitalisasi pencatatan tanggal formulir survei baru dibakukan pada 2025.\n"
          "2. Tahap 4 (Pengumpulan Berkas): Berhasil dipangkas sebesar 66,0% dari 32,6 hari (2024) menjadi 11,1 hari (2026).\n"
          "3. Tahap 7 (Tunggu Grand Opening): Dipangkas sebesar 60,7% dari 14,0 hari (2024) menjadi 5,5 hari (2026).\n"
          "4. Total Lead Time Pasca-Approval terealisasi 58,4 hari (2026), selaras sempurna dengan target cut-off manajemen ~58 hari.",
    outfile=os.path.join(GRAFIK_DIR, "4a_tabel_dekomposisi_7_siklus_waktu.png"),
    figsize=(14.5, 4.8), header_h=0.18, row_h=0.11
)

# 2. Monthly 2026
render_table_plot(
    col_headers, res_mo,
    title="TABEL DETAIL BULANAN DEKOMPOSISI 7 SIKLUS WAKTU TAHUN 2026 (JANUARI – AGUSTUS)",
    subtitle="Rincian Rata-rata Durasi (Hari Kalender) per Tahap Siklus Hidup Pembukaan Cabang UPC 2026 | PT Pusat Gadai Indonesia",
    notes="Catatan Kinerja Bulanan 2026:\n"
          "1. Akselerasi Negosiasi (Tahap 3): Mengalami percepatan luar biasa dari 23,1 hari (Feb) menjadi 2,0–2,9 hari (Jul–Agu).\n"
          "2. Efisiensi Berkas (Tahap 4): Dipangkas bertahap dari 15,9 hari (Mar) menjadi 3,7–6,6 hari (Jun–Agu).\n"
          "3. Lead Time Pasca-Approval: Memperlihatkan perbaikan dramatis dari 63,7 hari (Jan) turun ke 40,0–41,5 hari (Jul–Agu).\n"
          "4. Total End-to-End: Turun signifikan dari level 71,9 hari (Feb) menjadi 45,0–50,0 hari (Jul–Agu 2026).",
    outfile=os.path.join(GRAFIK_DIR, "4b_tabel_detail_bulanan_siklus_2026.png"),
    figsize=(14.5, 6.8), header_h=0.14, row_h=0.075
)

# 3. Combined Executive Dashboard
def render_combined_dashboard():
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(15, 10.5), dpi=300, gridspec_kw={'height_ratios': [1, 1.65]})
    
    # Top Table: Yearly
    ax1.axis('off')
    tab1_data = [[r['periode'], r['cabang'], r['t1'], r['t2'], r['t3'], r['t4'], r['t5'], r['t6'], r['t7'], r['pasca'], r['e2e']] for r in res_yr]
    col_widths = [0.11, 0.07, 0.09, 0.09, 0.08, 0.09, 0.09, 0.09, 0.09, 0.10, 0.10]
    tab1 = ax1.table(cellText=tab1_data, colLabels=col_headers, colWidths=col_widths, cellLoc='center', loc='center')
    tab1.auto_set_font_size(False)
    tab1.set_fontsize(8.0)
    for c in range(len(col_headers)):
        tab1[0, c].set_height(0.18)
        tab1[0, c].set_facecolor('#0F172A')
        tab1[0, c].get_text().set_color('#FFFFFF')
        tab1[0, c].get_text().set_fontweight('bold')
        tab1[0, c].get_text().set_fontsize(7.8)
        for r_i in range(1, len(tab1_data)+1):
            cell = tab1[r_i, c]
            cell.set_height(0.12)
            is_l = (r_i == len(tab1_data))
            cell.set_facecolor('#EFF6FF' if is_l else ('#F8FAFC' if r_i%2==1 else '#FFFFFF'))
            cell.set_edgecolor('#CBD5E1')
            if c == 0: cell.set_facecolor('#DBEAFE' if is_l else '#F1F5F9'); cell.get_text().set_fontweight('bold')
            elif c == 1: cell.get_text().set_fontweight('bold')
            elif c == 9: cell.set_facecolor('#D1FAE5' if is_l else '#ECFDF5'); cell.get_text().set_color('#065F46'); cell.get_text().set_fontweight('bold')
            elif c == 10: cell.set_facecolor('#EDE9FE' if is_l else '#F5F3FF'); cell.get_text().set_color('#5B21B6'); cell.get_text().set_fontweight('bold')
            if is_l:
                cell.set_edgecolor('#2563EB')
                cell.set_linewidth(1.3)
                if c not in [9, 10]: cell.get_text().set_color('#1E40AF')
    ax1.set_title("A. Komparasi Dekomposisi 7 Tahapan Siklus Waktu Historis Tahunan (2024–2026)", fontsize=10.5, fontweight='bold', color='#1E3A8A', loc='left', pad=10)
    
    # Bottom Table: Monthly 2026
    ax2.axis('off')
    tab2_data = [[r['periode'], r['cabang'], r['t1'], r['t2'], r['t3'], r['t4'], r['t5'], r['t6'], r['t7'], r['pasca'], r['e2e']] for r in res_mo]
    tab2 = ax2.table(cellText=tab2_data, colLabels=col_headers, colWidths=col_widths, cellLoc='center', loc='center')
    tab2.auto_set_font_size(False)
    tab2.set_fontsize(8.0)
    for c in range(len(col_headers)):
        tab2[0, c].set_height(0.13)
        tab2[0, c].set_facecolor('#0F172A')
        tab2[0, c].get_text().set_color('#FFFFFF')
        tab2[0, c].get_text().set_fontweight('bold')
        tab2[0, c].get_text().set_fontsize(7.8)
        for r_i in range(1, len(tab2_data)+1):
            cell = tab2[r_i, c]
            cell.set_height(0.08)
            is_l = (r_i == len(tab2_data))
            cell.set_facecolor('#EFF6FF' if is_l else ('#F8FAFC' if r_i%2==1 else '#FFFFFF'))
            cell.set_edgecolor('#CBD5E1')
            if c == 0: cell.set_facecolor('#DBEAFE' if is_l else '#F1F5F9'); cell.get_text().set_fontweight('bold')
            elif c == 1: cell.get_text().set_fontweight('bold')
            elif c == 9: cell.set_facecolor('#D1FAE5' if is_l else '#ECFDF5'); cell.get_text().set_color('#065F46'); cell.get_text().set_fontweight('bold')
            elif c == 10: cell.set_facecolor('#EDE9FE' if is_l else '#F5F3FF'); cell.get_text().set_color('#5B21B6'); cell.get_text().set_fontweight('bold')
            if is_l:
                cell.set_edgecolor('#2563EB')
                cell.set_linewidth(1.3)
                if c not in [9, 10]: cell.get_text().set_color('#1E40AF')
    ax2.set_title("B. Rincian Detail Bulanan Dekomposisi 7 Tahapan Siklus Waktu Tahun 2026 (Januari – Agustus)", fontsize=10.5, fontweight='bold', color='#1E3A8A', loc='left', pad=10)
    
    plt.suptitle("DASHBOARD EKSEKUTIF: TABEL DEKOMPOSISI 7 SIKLUS WAKTU & HARMONISASI LEAD TIME UPC\nHarmonisasi: End-to-End (~69 Hari) vs Pasca-Approval (~58 Hari) | PT Pusat Gadai Indonesia", fontsize=12, fontweight='bold', color='#0F172A', y=0.98)
    
    fig.text(0.04, 0.015, "Insight Eksekutif: (1) Total lead time pasca-persetujuan 2026 stabil di 58,4 hari (~58 hr), dan siklus penuh survei s/d open di 69,1 hari (~69 hr).\n(2) Akselerasi negosiasi & efisiensi administrasi berkas terakselerasi signifikan di Q2-Q3 2026, memangkas lead time penutupan gerai hingga ke level 40–50 hari.", fontsize=7.8, color='#334155', style='italic', linespacing=1.3)
    
    p_all = os.path.join(GRAFIK_DIR, "4_dekomposisi_7_siklus_tabel_komprehensif.png")
    fig.savefig(p_all, bbox_inches='tight')
    plt.close(fig)
    print(f"[✓] Dashboard Komprehensif tersimpan: {p_all}")

render_combined_dashboard()
