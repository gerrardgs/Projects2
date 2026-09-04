import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# ============================================================================
# 0. PATH SETUP
# ============================================================================
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))           # .../train
DATASET_PATH = os.path.join(BASE_DIR, 'trainfile', 'dataset_final.csv')
OUT_DIR      = os.path.join(BASE_DIR, 'paper_assets')
OUT_PATH     = os.path.join(OUT_DIR, 'figure8_label_distribution.png')
os.makedirs(OUT_DIR, exist_ok=True)

# ============================================================================
# 1. LOAD DATA (defensive: validasi file & kolom)
# ============================================================================
if not os.path.exists(DATASET_PATH):
    raise FileNotFoundError(
        f"dataset_final.csv tidak ditemukan di {DATASET_PATH}.\n"
        f"Jalankan fuzzytrainv2.py dulu untuk menghasilkan dataset final."
    )

df = pd.read_csv(DATASET_PATH)

required_cols = ['label_who_ukur']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise KeyError(
        f"Kolom yang dibutuhkan tidak ada di dataset_final.csv: {missing}\n"
        f"Kolom yang tersedia: {list(df.columns)}"
    )

df = df.dropna(subset=required_cols).reset_index(drop=True)
n_total = len(df)

# ============================================================================
# 2. HITUNG DISTRIBUSI LABEL
# ============================================================================
labels = ['Normal', 'Stunting']

counts = []
for lbl in labels:
    n_sub = int((df['label_who_ukur'] == lbl).sum())
    counts.append(n_sub)

# Sanity check: pastikan jumlah baris ter-cover semua label
covered = sum(counts)
if covered != n_total:
    unknown = df[~df['label_who_ukur'].isin(labels)]['label_who_ukur'].unique().tolist()
    print(f"[!] Peringatan: {n_total - covered} baris memiliki label "
          f"yang tidak dikenal: {unknown}")

percentages = [c / n_total * 100 if n_total > 0 else 0.0 for c in counts]

# ============================================================================
# 3. PLOT
# ============================================================================
plt.rcParams.update({
    'font.family':     'serif',
    'font.size':       11,
    'axes.titlesize':  12,
    'axes.labelsize':  11,
    'legend.fontsize': 10,
    'xtick.labelsize': 11,
    'ytick.labelsize': 10,
    'axes.linewidth':  1.0,
})

fig, ax = plt.subplots(figsize=(6.0, 4.6))

# Warna: Normal = biru kalem, Stunting = merah kalem (konsisten dengan figure 4 & 5)
colors = ['tab:blue', 'tab:red']

bars = ax.bar(labels, counts,
              color=colors, edgecolor='black', linewidth=1.0,
              width=0.55)

# Anotasi di atas tiap bar: count + persentase
for bar, n_sub, pct in zip(bars, counts, percentages):
    if n_sub == 0:
        ax.text(bar.get_x() + bar.get_width() / 2, 0.5,
                'n=0', ha='center', va='bottom',
                fontsize=10, fontweight='bold', color='gray')
        continue
    ax.text(bar.get_x() + bar.get_width() / 2,
            bar.get_height() + n_total * 0.012,
            f'{n_sub}\n({pct:.1f}%)',
            ha='center', va='bottom',
            fontsize=11, fontweight='bold')

ax.set_ylabel('Jumlah Balita')
ax.set_xlabel('Label WHO Z-Score (Saat Pengukuran)')
ax.set_title(f'Distribusi Label Z-Score Saat Pengukuran (n = {n_total})')
ax.grid(True, alpha=0.3, axis='y', linestyle='--')

# Y-axis: cukup ruang untuk anotasi di atas bar tertinggi
y_max = max(counts) if max(counts) > 0 else 1
ax.set_ylim(0, y_max * 1.18)

plt.tight_layout()
plt.savefig(OUT_PATH, dpi=300, bbox_inches='tight')
print(f"[OK] Figure tersimpan: {OUT_PATH}")

# ============================================================================
# 4. PRINT STATISTIK (untuk caption paper / diskusi pembimbing)
# ============================================================================
print(f"\n--- Distribusi Label Z-Score Saat Pengukuran ---")
print(f"  Total samples : {n_total}")
for lbl, n_sub, pct in zip(labels, counts, percentages):
    print(f"  {lbl:10s}: {n_sub:>4d}  ({pct:>5.2f}%)")

plt.show()