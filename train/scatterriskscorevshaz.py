import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import spearmanr

# ============================================================================
# 0. PATH SETUP
# ============================================================================
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))           # .../train
DATASET_PATH = os.path.join(BASE_DIR, 'trainfile', 'dataset_final.csv')
OUT_DIR      = os.path.join(BASE_DIR, 'paper_assets')
OUT_PATH     = os.path.join(OUT_DIR, 'figure5_scatter_risk_haz.png')
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

required_cols = ['risk_score', 'z_score_ukur', 'label_who_ukur']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise KeyError(
        f"Kolom yang dibutuhkan tidak ada di dataset_final.csv: {missing}\n"
        f"Kolom yang tersedia: {list(df.columns)}"
    )

# Dropna eksplisit pada kolom yang dipakai → n di title pasti benar
df = df.dropna(subset=required_cols).reset_index(drop=True)
n = len(df)

# ============================================================================
# 2. KORELASI SPEARMAN
# ============================================================================
# rho, p_val = spearmanr(df['risk_score'], df['z_score_ukur'])

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

fig, ax = plt.subplots(figsize=(7.2, 5.0))

# Warna per kelas (Normal=biru, Stunting=merah) — konsisten dengan figure 4
mask_stunting = df['label_who_ukur'] == 'Stunting'
mask_normal   = ~mask_stunting

ax.scatter(df.loc[mask_normal,   'risk_score'],
           df.loc[mask_normal,   'z_score_ukur'],
           c='tab:blue', alpha=0.6, s=32,
           edgecolor='white', linewidth=0.5,
           label=f'Normal (n={mask_normal.sum()})')

ax.scatter(df.loc[mask_stunting, 'risk_score'],
           df.loc[mask_stunting, 'z_score_ukur'],
           c='tab:red', alpha=0.7, s=32,
           edgecolor='white', linewidth=0.5,
           label=f'Stunting (n={mask_stunting.sum()})')

# Threshold WHO HAZ = -2 (batas stunting)
ax.axhline(y=-2, color='black', linestyle='--', alpha=0.55, linewidth=1.2,
           label='Stunting threshold (HAZ = −2)')

# Garis regresi linear (visual trend, bukan claim kausalitas)
slope, intercept = np.polyfit(df['risk_score'], df['z_score_ukur'], 1)
xr = np.linspace(df['risk_score'].min(), df['risk_score'].max(), 100)
ax.plot(xr, slope * xr + intercept,
        color='gray', linewidth=2.0, alpha=0.7,
        label='Linear fit')

# # Anotasi statistik (kiri-bawah)
# p_str = '< 0.001' if p_val < 0.001 else f'= {p_val:.3f}'
# ax.text(0.04, 0.06,
#         f'Spearman ρ = {rho:.3f}  (p {p_str})',
#         transform=ax.transAxes, fontsize=11,
#         bbox=dict(boxstyle='round,pad=0.45',
#                   facecolor='white', edgecolor='gray', alpha=0.92))

ax.set_xlabel('Fuzzy Risk Score')
ax.set_ylabel('WHO HAZ at measurement')
ax.set_title(f'Risk Score vs HAZ (n = {n})')
ax.legend(loc='upper right', framealpha=0.92, edgecolor='gray')
ax.grid(True, alpha=0.3, linestyle='--')

plt.tight_layout()
plt.savefig(OUT_PATH, dpi=300, bbox_inches='tight')
print(f"[OK] Figure tersimpan: {OUT_PATH}")

# ============================================================================
# 4. PRINT STATISTIK (untuk caption paper / diskusi pembimbing)
# ============================================================================
print(f"\n--- Dataset ---")
print(f"  Total samples plotted : {n}")
print(f"  Normal                : {mask_normal.sum()}")
print(f"  Stunting              : {mask_stunting.sum()}")

print(f"\n--- Risk Score Statistics ---")
rs_n = df.loc[mask_normal,   'risk_score']
rs_s = df.loc[mask_stunting, 'risk_score']
print(f"  Normal   : mean={rs_n.mean():6.2f}  median={rs_n.median():6.2f}  "
      f"min={rs_n.min():6.2f}  max={rs_n.max():6.2f}")
print(f"  Stunting : mean={rs_s.mean():6.2f}  median={rs_s.median():6.2f}  "
      f"min={rs_s.min():6.2f}  max={rs_s.max():6.2f}")

# print(f"\n--- Spearman Correlation (Risk Score vs HAZ) ---")
# print(f"  rho     : {rho:.4f}")
# print(f"  p-value : {p_val:.6f}")
# print(f"  Linear  : HAZ = {slope:.4f} * RiskScore + {intercept:.4f}")

plt.show()