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
OUT_PATH     = os.path.join(OUT_DIR, 'figure6_prevalence_stratification.png')
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

required_cols = ['priority_flag', 'label_who_ukur']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise KeyError(
        f"Kolom yang dibutuhkan tidak ada di dataset_final.csv: {missing}"
    )

df = df.dropna(subset=required_cols).reset_index(drop=True)
n_total = len(df)

# ============================================================================
# 2. HITUNG STUNTING RATE PER PRIORITY FLAG
# ============================================================================
# Flag values harus persis sesuai output get_priority_flag() di fuzzytrainv2.py
flags        = ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']
labels_short = ['Normal', 'Low',         'Medium',          'High']

baseline = (df['label_who_ukur'] == 'Stunting').mean()

rates  = []
counts = []
n_stunting_per_flag = []

for f in flags:
    subset = df[df['priority_flag'] == f]
    n_sub = len(subset)
    n_stu = int((subset['label_who_ukur'] == 'Stunting').sum())
    rate = n_stu / n_sub if n_sub > 0 else 0.0
    rates.append(rate)
    counts.append(n_sub)
    n_stunting_per_flag.append(n_stu)

# Sanity check: pastikan jumlah baris ter-cover semua flag (tidak ada flag tak terduga)
covered = sum(counts)
if covered != n_total:
    unknown = df[~df['priority_flag'].isin(flags)]['priority_flag'].unique().tolist()
    print(f"[!] Peringatan: {n_total - covered} baris memiliki priority_flag "
          f"yang tidak dikenal: {unknown}")

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

# Warna gradien risiko (hijau → kuning → oranye → merah)
colors = ['#2ca02c', '#f5c518', '#ff7f0e', '#d62728']

bars = ax.bar(labels_short, rates,
              color=colors, edgecolor='black', linewidth=1.0,
              width=0.62)

# Baseline line
ax.axhline(y=baseline, color='steelblue', linestyle='--', linewidth=2.0,
           label=f'Baseline prevalence ({baseline*100:.1f}%)', zorder=3)

# Anotasi di atas tiap bar: lift ratio + n
for bar, rate, n_sub, n_stu in zip(bars, rates, counts, n_stunting_per_flag):
    if n_sub == 0:
        # Bar kosong: tampilkan "n=0" supaya jelas tidak ada data, bukan rate=0 yang bias
        ax.text(bar.get_x() + bar.get_width() / 2, 0.005,
                'n=0', ha='center', va='bottom',
                fontsize=10, fontweight='bold', color='gray')
        continue
    ratio = rate / baseline if baseline > 0 else 0.0
    ax.text(bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.012,
            f'{ratio:.2f}×\n({n_stu}/{n_sub})',
            ha='center', va='bottom',
            fontsize=10, fontweight='bold')

ax.set_ylabel('Stunting Rate')
ax.set_xlabel('Priority Flag')
ax.set_title(f'Empirical Stunting Rate Stratified by Priority Flag (n = {n_total})')
ax.legend(loc='upper left', framealpha=0.92, edgecolor='gray')
ax.grid(True, alpha=0.3, axis='y', linestyle='--')

# Y-axis: cukup ruang untuk anotasi di atas bar tertinggi
y_max = max(rates) if max(rates) > 0 else baseline
ax.set_ylim(0, y_max * 1.32 if y_max > 0 else 0.1)

# Format y-axis sebagai persen untuk readability paper
from matplotlib.ticker import FuncFormatter
ax.yaxis.set_major_formatter(FuncFormatter(lambda y, _: f'{y*100:.0f}%'))

plt.tight_layout()
plt.savefig(OUT_PATH, dpi=300, bbox_inches='tight')
print(f"[OK] Figure tersimpan: {OUT_PATH}")

# ============================================================================
# 4. PRINT TABEL STRATIFIKASI (untuk caption / tabel paper)
# ============================================================================
print(f"\n--- Dataset ---")
print(f"  Total samples            : {n_total}")
print(f"  Total stunting (gold)    : {int((df['label_who_ukur'] == 'Stunting').sum())}")
print(f"  Baseline prevalence      : {baseline*100:.2f}%")

print(f"\n--- Stunting Rate per Priority Flag ---")
print(f"  {'Flag':18s} {'n':>5s} {'Stunting':>9s} {'Rate':>8s} {'Lift':>7s}")
print(f"  {'-'*18} {'-'*5} {'-'*9} {'-'*8} {'-'*7}")
for f, lbl, n_sub, n_stu, rate in zip(flags, labels_short, counts,
                                        n_stunting_per_flag, rates):
    lift = rate / baseline if baseline > 0 else 0.0
    rate_str = f'{rate*100:.2f}%' if n_sub > 0 else '   n/a '
    lift_str = f'{lift:.2f}x' if n_sub > 0 else '  n/a '
    print(f"  {f:18s} {n_sub:>5d} {n_stu:>9d} {rate_str:>8s} {lift_str:>7s}")

plt.show()