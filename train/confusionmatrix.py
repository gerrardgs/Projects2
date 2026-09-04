import os
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ============================================================================
# 0. PATH SETUP
# ============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))          # .../train
OUT_DIR  = os.path.join(BASE_DIR, 'paper_assets')
OUT_PATH = os.path.join(OUT_DIR, 'figure4_confusion_matrix.png')
os.makedirs(OUT_DIR, exist_ok=True)

# ============================================================================
# 1. CONFUSION MATRIX (hasil aktual GA-optimized model di test set)
# ============================================================================
#                     Predicted Normal   Predicted Stunting
# Actual Normal   |        TN=46                FP=19
# Actual Stunting |        FN= 2                TP=11
cm = np.array([
    [46, 19],
    [ 2, 11],
])

TN, FP = cm[0]
FN, TP = cm[1]
n_total = TN + FP + FN + TP

# ============================================================================
# 2. METRIK (untuk anotasi di luar plot bila perlu, sekaligus verifikasi)
# ============================================================================
acc  = (TP + TN) / n_total
prec = TP / (TP + FP) if (TP + FP) > 0 else 0.0
rec  = TP / (TP + FN) if (TP + FN) > 0 else 0.0
spec = TN / (TN + FP) if (TN + FP) > 0 else 0.0
f1   = 2 * prec * rec / (prec + rec) if (prec + rec) > 0 else 0.0

# ============================================================================
# 3. PLOT
# ============================================================================
plt.rcParams.update({
    'font.family':     'serif',
    'font.size':       11,
    'axes.titlesize':  12,
    'axes.labelsize':  11,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'axes.linewidth':  1.0,
})

fig, ax = plt.subplots(figsize=(5.2, 4.2))

sns.heatmap(
    cm,
    annot=True,
    fmt='d',
    cmap='Blues',
    cbar=False,
    xticklabels=['Normal', 'Stunting'],
    yticklabels=['Normal', 'Stunting'],
    annot_kws={'size': 18, 'weight': 'bold'},
    linewidths=1.0,
    linecolor='white',
    ax=ax,
)

ax.set_xlabel('Predicted', labelpad=8)
ax.set_ylabel('Actual (WHO HAZ)', labelpad=8)
ax.set_title(f'Confusion Matrix — GA-Optimized Model on Test Set (n={n_total})',
             pad=10)

# Pastikan label sumbu tidak terpotong & rapi
plt.setp(ax.get_yticklabels(), rotation=0)
plt.setp(ax.get_xticklabels(), rotation=0)

plt.tight_layout()
plt.savefig(OUT_PATH, dpi=300, bbox_inches='tight')
print(f"[OK] Figure tersimpan: {OUT_PATH}")

# ============================================================================
# 4. PRINT METRIK (untuk caption paper / diskusi pembimbing)
# ============================================================================
print("\n--- Confusion Matrix ---")
print(f"  TN = {TN:>3d}    FP = {FP:>3d}")
print(f"  FN = {FN:>3d}    TP = {TP:>3d}")
print(f"\n  Total test set: {n_total}")
print(f"  Actual Normal   : {TN + FP}")
print(f"  Actual Stunting : {FN + TP}")

print("\n--- Performance Metrics ---")
print(f"  Accuracy    : {acc:.4f}  ({acc*100:.2f}%)")
print(f"  Precision   : {prec:.4f}  ({prec*100:.2f}%)")
print(f"  Recall/Sens : {rec:.4f}  ({rec*100:.2f}%)")
print(f"  Specificity : {spec:.4f}  ({spec*100:.2f}%)")
print(f"  F1-Score    : {f1:.4f}")

plt.show()