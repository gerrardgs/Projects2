# TAMBAHKAN ke fuzzytrainv2.py setelah Task 5.4, atau jalankan terpisah:
from scipy.stats import spearmanr
import pandas as pd

df = pd.read_csv('train/trainfile/dataset_final.csv')

# 1. Spearman correlation
rho, p_val = spearmanr(df['risk_score'], df['z_score_ukur'])
print(f"Spearman ρ = {rho:.4f}, p-value = {p_val:.4e}")

# 2. Prevalence stratification
baseline = (df['label_who_ukur'] == 'Stunting').mean()
print(f"\nBaseline stunting rate: {baseline:.4f} ({baseline*100:.2f}%)")

print("\n=== Prevalence per priority flag ===")
for flag in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']:
    subset = df[df['priority_flag'] == flag]
    if len(subset) == 0:
        print(f"{flag}: n=0")
        continue
    n_stunt = (subset['label_who_ukur'] == 'Stunting').sum()
    rate = n_stunt / len(subset)
    ratio = rate / baseline if baseline > 0 else 0
    print(f"{flag:18s}: n={len(subset):3d}, stunt={n_stunt:3d}, rate={rate:.4f} ({rate*100:.1f}%), ratio={ratio:.2f}x")

# 3. Pearson correlations of inputs with HAZ
print("\n=== Input vs HAZ correlations ===")
for col in ['pbl_cm', 'tinggi_badan_cm', 'berat_badan_gram', 'mpasi_bulan']:
    pearson_r = df[col].corr(df['z_score_ukur'])
    spearman_r = df[col].corr(df['z_score_ukur'], method='spearman')
    print(f"{col:25s}: Pearson r={pearson_r:+.4f}, Spearman ρ={spearman_r:+.4f}")
    