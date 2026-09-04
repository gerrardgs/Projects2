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
OUT_PATH     = os.path.join(OUT_DIR, 'figure10_growth_trajectory.png')
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

required_cols = ['label_who', 'fuzzy_label']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise KeyError(
        f"Kolom yang dibutuhkan tidak ada di dataset_final.csv: {missing}\n"
        f"Kolom yang tersedia: {list(df.columns)}"
    )

df = df.dropna(subset=required_cols).reset_index(drop=True)
n_total = len(df)

# ============================================================================
# 2. HITUNG TRAYEKTORI PERTUMBUHAN
# ============================================================================
# Definisi 4 trayektori dari kombinasi (label_lahir, label_ukur)
# Order: dari paling ringan (konsisten sehat) ke paling berat (risiko persisten)
trajectories = [
    ('Normal -> Normal',     'Normal',   'Normal',   'Konsisten sehat'),
    ('Stunting -> Normal',   'Stunting', 'Normal',   'Pemulihan positif'),
    ('Normal -> Stunting',   'Normal',   'Stunting', 'Trayektori negatif'),
    ('Stunting -> Stunting', 'Stunting', 'Stunting', 'Risiko persisten'),
]

labels_short = ['Konsisten\nSehat', 'Pemulihan\nPositif',
                'Trayektori\nNegatif', 'Risiko\nPersisten']

counts = []
for _, lbl_lahir, lbl_ukur, _ in trajectories:
    mask = (df['label_who'] == lbl_lahir) & (df['fuzzy_label'] == lbl_ukur)
    n_sub = int(mask.sum())
    counts.append(n_sub)

# Sanity check: pastikan jumlah baris ter-cover semua trayektori
covered = sum(counts)
if covered != n_total:
    unknown_pairs = (
        df[~df.apply(lambda r: any(
            (r['label_who'] == t[1]) and (r['fuzzy_label'] == t[2])
            for t in trajectories
        ), axis=1)][['label_who', 'fuzzy_label']]
        .drop_duplicates()
        .values.tolist()
    )
    print(f"[!] Peringatan: {n_total - covered} baris memiliki kombinasi label "
          f"yang tidak dikenal: {unknown_pairs}")

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
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'axes.linewidth':  1.0,
})

fig, ax = plt.subplots(figsize=(8.0, 5.2))

# Warna gradien semantik:
# Konsisten Sehat = hijau, Pemulihan Positif = biru,
# Trayektori Negatif = oranye, Risiko Persisten = merah
colors = ['#2ca02c', '#1f77b4', '#ff7f0e', '#d62728']

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
ax.set_xlabel('Trayektori Pertumbuhan (Label Lahir -> Output Sistem Fuzzy)')
ax.set_title(f'Trayektori Pertumbuhan dari Lahir ke Output Sistem Fuzzy (n = {n_total})')
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
print(f"\n--- Trayektori Pertumbuhan Lahir -> Output Sistem Fuzzy (Post-GA) ---")
print(f"  Total samples : {n_total}")
for (label_full, _, _, interp), n_sub, pct in zip(trajectories, counts, percentages):
    print(f"  {label_full:24s}: {n_sub:>4d}  ({pct:>5.2f}%)  -> {interp}")

# Ringkasan operasional: total balita yang berpotensi diintervensi
n_intervention = counts[2] + counts[3]   # Trayektori Negatif + Risiko Persisten
pct_intervention = n_intervention / n_total * 100 if n_total > 0 else 0.0
print(f"\n--- Ringkasan Operasional ---")
print(f"  Stunting saat ukur (target intervensi): {n_intervention} ({pct_intervention:.2f}%)")
print(f"  Trayektori negatif (target prioritas) : {counts[2]} ({percentages[2]:.2f}%)")

plt.show()