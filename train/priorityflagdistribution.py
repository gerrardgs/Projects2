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
OUT_PATH     = os.path.join(OUT_DIR, 'figure9_priority_flag_distribution.png')
os.makedirs(OUT_DIR, exist_ok=True)

# ============================================================================
# 1. LOAD DATA (defensive)
# ============================================================================
if not os.path.exists(DATASET_PATH):
    raise FileNotFoundError(
        f"dataset_final.csv tidak ditemukan di {DATASET_PATH}.\n"
        f"Jalankan fuzzytrainv2.py dulu untuk menghasilkan dataset final."
    )

df = pd.read_csv(DATASET_PATH)

required_cols = ['priority_flag']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise KeyError(
        f"Kolom yang dibutuhkan tidak ada di dataset_final.csv: {missing}"
    )

df = df.dropna(subset=required_cols).reset_index(drop=True)
n_total = len(df)

# ============================================================================
# 2. HITUNG DISTRIBUSI PRIORITY FLAG
# ============================================================================
# Flag values harus persis sesuai output get_priority_flag() di fuzzytrainv2.py
flags        = ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']
labels_short = ['Normal', 'Low',          'Medium',          'High']

counts = []
for f in flags:
    n_sub = int((df['priority_flag'] == f).sum())
    counts.append(n_sub)

# Sanity check: pastikan jumlah baris ter-cover semua flag
covered = sum(counts)
if covered != n_total:
    unknown = df[~df['priority_flag'].isin(flags)]['priority_flag'].unique().tolist()
    print(f"[!] Peringatan: {n_total - covered} baris memiliki priority_flag "
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

fig, ax = plt.subplots(figsize=(7.2, 5.0))

# Warna gradien risiko (hijau -> kuning -> oranye -> merah)
# Konsisten dengan stuntingrateperpriorflag.py
colors = ['#2ca02c', '#f5c518', '#ff7f0e', '#d62728']

bars = ax.bar(labels_short, counts,
              color=colors, edgecolor='black', linewidth=1.0,
              width=0.62)

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
            fontsize=10, fontweight='bold')

ax.set_ylabel('Jumlah Balita')
ax.set_xlabel('Priority Flag')
ax.set_title(f'Distribusi Priority Flag Sistem Stuntor (n = {n_total})')
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
print(f"\n--- Distribusi Priority Flag ---")
print(f"  Total samples : {n_total}")
for f, n_sub, pct in zip(flags, counts, percentages):
    print(f"  {f:18s}: {n_sub:>4d}  ({pct:>5.2f}%)")

plt.show()