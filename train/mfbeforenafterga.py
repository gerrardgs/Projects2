import os
import sys
import json
import numpy as np
import matplotlib.pyplot as plt

# ============================================================================
# 0. PATH SETUP (resolusi relatif terhadap lokasi file ini, bukan CWD)
# ============================================================================
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))            # .../train
PROJECT_DIR = os.path.dirname(BASE_DIR)                            # .../Projects2
MODEL_PATH = os.path.join(BASE_DIR, 'trainfile', 'trained_model.json')
# paper_assets/ berada di dalam train/ (sesuai struktur project Mas Ali)
OUT_DIR    = os.path.join(BASE_DIR, 'paper_assets')
OUT_PATH   = os.path.join(OUT_DIR, 'figure3_mf_male_height.png')

# Pastikan paper_assets/ ada (FIX utama untuk error FileNotFoundError)
os.makedirs(OUT_DIR, exist_ok=True)

# Import baseline dari source-of-truth supaya konsisten dengan training
sys.path.insert(0, BASE_DIR)        # cari fuzzytrainv2 di folder train/
try:
    from fuzzytrainv2 import TB_L_MF_INIT
except ImportError:
    # Fallback bila fuzzytrainv2.py ada di project root, bukan di train/
    sys.path.insert(0, PROJECT_DIR)
    from fuzzytrainv2 import TB_L_MF_INIT

# ============================================================================
# 1. LOAD POST-GA MF DARI trained_model.json
# ============================================================================
if not os.path.exists(MODEL_PATH):
    raise FileNotFoundError(
        f"trained_model.json tidak ditemukan di {MODEL_PATH}.\n"
        f"Jalankan fuzzytrainv2.py dulu untuk menghasilkan model."
    )

with open(MODEL_PATH, 'r') as f:
    model = json.load(f)

post_mf_raw = model['mf_params']['tb_l']    # boys

# Format JSON: {"Pendek": ["trap", [a,b,c,d]], ...}
# Konversi ke {"Pendek": ("trap", [a,b,c,d]), ...}
post_mf = {name: (val[0], val[1]) for name, val in post_mf_raw.items()}

# init_mf langsung dari source-of-truth (sudah berbentuk tuple)
init_mf = {name: (val[0], list(val[1])) for name, val in TB_L_MF_INIT.items()}

# ============================================================================
# 2. FUNGSI KEANGGOTAAN (mendukung trap & tri)
# ============================================================================
def trapezoid(x, a, b, c, d):
    if x <= a or x >= d:
        return 0.0
    if a < x < b:
        return (x - a) / (b - a) if b != a else 1.0
    if b <= x <= c:
        return 1.0
    if c < x < d:
        return (d - x) / (d - c) if d != c else 1.0
    return 0.0

def triangle(x, a, b, c):
    if x <= a or x >= c:
        return 0.0
    if a < x <= b:
        return (x - a) / (b - a) if b != a else 1.0
    if b < x < c:
        return (c - x) / (c - b) if c != b else 1.0
    return 0.0

def eval_mf(x, mf_type, params):
    if mf_type == 'trap':
        return trapezoid(x, *params)
    if mf_type == 'tri':
        return triangle(x, *params)
    raise ValueError(f"MF type tidak dikenal: {mf_type}")

# ============================================================================
# 3. PLOT
# ============================================================================
# Range x cukup untuk menampilkan semua kategori (Pendek mulai 38)
x_range = np.linspace(38, 130, 600)

# Konfigurasi visual akademik (paper-grade)
plt.rcParams.update({
    'font.family':      'serif',
    'font.size':        11,
    'axes.titlesize':   12,
    'axes.labelsize':   11,
    'legend.fontsize':  10,
    'xtick.labelsize':  10,
    'ytick.labelsize':  10,
    'axes.linewidth':   1.0,
})

fig, axes = plt.subplots(1, 2, figsize=(12, 4.2), sharey=True)
colors = {'Pendek': '#d62728', 'Normal': '#2ca02c', 'Tinggi': '#1f77b4'}
linestyles = {'Pendek': '-', 'Normal': '-', 'Tinggi': '-'}
order = ['Pendek', 'Normal', 'Tinggi']  # urutan legend konsisten

def plot_mf(ax, mf_dict, title):
    for name in order:
        if name not in mf_dict:
            continue
        mf_type, params = mf_dict[name]
        y = np.array([eval_mf(x, mf_type, params) for x in x_range])
        ax.plot(x_range, y,
                label=name,
                color=colors[name],
                linestyle=linestyles[name],
                linewidth=2.0)
    ax.set_xlabel('Height (cm)')
    ax.set_title(title)
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.set_ylim(-0.02, 1.08)
    ax.set_xlim(38, 130)
    ax.set_yticks([0, 0.25, 0.5, 0.75, 1.0])
    ax.legend(loc='upper right', framealpha=0.92, edgecolor='gray')

axes[0].set_ylabel('Membership degree μ')
plot_mf(axes[0], init_mf, '(a) Before GA Optimization')
plot_mf(axes[1], post_mf, '(b) After GA Optimization')

plt.tight_layout()
plt.savefig(OUT_PATH, dpi=300, bbox_inches='tight')
print(f"[OK] Figure tersimpan: {OUT_PATH}")

# Tampilkan angka boundary untuk verifikasi (memudahkan diskusi pembimbing)
print("\n--- Boundary Points (Before GA) ---")
for n in order:
    print(f"  {n:7s} : {init_mf[n][0]:4s} {init_mf[n][1]}")
print("\n--- Boundary Points (After GA)  ---")
for n in order:
    t, p = post_mf[n]
    p_round = [round(v, 3) for v in p]
    print(f"  {n:7s} : {t:4s} {p_round}")

plt.show()